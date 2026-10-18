using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus.Internal;
using Xunit;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

namespace Docxodus.Tests;

public class WmlToHtmlCssSerializationTests
{
    private static byte[] Fixture()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(new Wp.Body(
                new Wp.Paragraph(new Wp.Run(new Wp.Text("Cedar cascade sample & control."))),
                new Wp.SectionProperties(new Wp.PageSize { Width = 12240, Height = 15840 })));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            doc.Save();
            Assert.Empty(new OpenXmlValidator().Validate(doc));
        }
        return stream.ToArray();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ChildSelectors_SurviveDefaultAndCompactSerialization(bool classes)
    {
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("cascade.docx", Fixture()),
            new WmlToHtmlConverterSettings
            {
                FabricateCssClasses = classes,
                GeneralCss = "p > span { color: rgb(255, 0, 0); }",
                AdditionalCss = "p > span { background-color: rgb(0, 255, 0); }",
            });
        foreach (var options in new[] { SaveOptions.None, SaveOptions.DisableFormatting })
        {
            var serialized = html.ToString(options);
            Assert.Contains("p > span { color:", serialized);
            Assert.Contains("p > span { background-color:", serialized);
            Assert.Equal("Cedar cascade sample & control.", XElement.Parse(serialized)
                .Descendants().Single(e => e.Name.LocalName == "p").Value);
        }
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void SharedFacade_PreservesSelectorsAndCssStringValues(int pagination)
    {
        var serialized = HtmlConversionOps.ConvertToHtml(Fixture(), new HtmlConversionOptions
        {
            PaginationMode = pagination,
            AdditionalCss = "p > span { color: red; } p::after { content: 'A&B > C'; }",
        });
        Assert.Contains("p > span { color: red; }", serialized);
        Assert.Contains("content: 'A&B > C'", serialized);
        Assert.NotNull(XElement.Parse(serialized));
    }
    [Fact]
    public void CssDelimiters_CannotEndTheStyleElementOrCdataSection()
    {
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("cascade.docx", Fixture()),
            new WmlToHtmlConverterSettings
            {
                AdditionalCss = "p::after { content: '</StYlE><p id=escape>A&B ]]> C'; }",
            });
        var serialized = html.ToString(SaveOptions.DisableFormatting);
        Assert.DoesNotContain("</StYlE>", serialized);
        Assert.DoesNotContain("&amp;B", serialized);
        Assert.NotNull(XElement.Parse(serialized));
    }

}
