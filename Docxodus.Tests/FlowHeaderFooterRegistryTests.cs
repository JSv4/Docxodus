using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Wp = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace Docxodus.Tests;

public class FlowHeaderFooterRegistryTests
{
    private static byte[] Document(bool first, bool even)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings(new Wp.EvenAndOddHeaders { Val = even });
            var section = new Wp.SectionProperties();
            foreach (var (type, name) in new[] { (Wp.HeaderFooterValues.Default, "default"),
                (Wp.HeaderFooterValues.First, "first"), (Wp.HeaderFooterValues.Even, "even") })
            {
                var header = main.AddNewPart<HeaderPart>();
                header.Header = new Wp.Header(new Wp.Paragraph(new Wp.Run(new Wp.Text("header-" + name))));
                var footer = main.AddNewPart<FooterPart>();
                footer.Footer = new Wp.Footer(new Wp.Paragraph(new Wp.Run(new Wp.Text("footer-" + name))));
                section.PrependChild(new Wp.HeaderReference { Type = type, Id = main.GetIdOfPart(header) });
                section.Append(new Wp.FooterReference { Type = type, Id = main.GetIdOfPart(footer) });
            }
            section.Append(new Wp.TitlePage { Val = first });
            main.Document = new Wp.Document(new Wp.Body(
                new Wp.Paragraph(new Wp.ParagraphProperties(section), new Wp.Run(new Wp.Text("Opening"))),
                new Wp.Paragraph(new Wp.Run(new Wp.Text("Following"))),
                new Wp.SectionProperties(new Wp.TitlePage { Val = first })));
        }
        return stream.ToArray();
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void FlowRegistryIncludesActiveInheritedVariants(bool first, bool even)
    {
        var bytes = Document(first, even);
        var html = XDocument.Parse(HtmlConversionOps.ConvertToHtml(bytes, new HtmlConversionOptions
        { RenderHeadersAndFooters = true, RenderHeaderFooterRegistry = true, StampAnchors = true }));
        var registry = Assert.Single(html.Descendants(), e => (string?)e.Attribute("id") == "document-hf-registry");
        Assert.Equal("display:none", (string?)registry.Attribute("style"));
        var types = new[] { "default" }.Concat(first ? new[] { "first" } : new string[0])
            .Concat(even ? new[] { "even" } : new string[0]);
        var expected = types.SelectMany(type => new[] { "header-" + type, "footer-" + type }).OrderBy(x => x).ToArray();
        foreach (var section in new[] { "0", "1" })
        {
            var entries = registry.Elements().Where(e => (string?)e.Attribute("data-section") == section).ToList();
            Assert.Equal(expected, entries.Select(e => (string)e.Attribute("data-hf-type")!).OrderBy(x => x));
            Assert.All(entries, e => Assert.Equal((string?)e.Attribute("data-hf-type"), e.Value));
        }
        Assert.Empty(html.Descendants(Xhtml.header));
        Assert.Empty(html.Descendants(Xhtml.footer));
        Assert.DoesNotContain(html.Descendants(), e => (string?)e.Attribute("id") == "pagination-staging");
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    public void RegistryAndStoryFlagsRetainTheirIndependentDefaults(bool renderStories, bool registry)
    {
        var html = XDocument.Parse(HtmlConversionOps.ConvertToHtml(Document(true, true), new HtmlConversionOptions
        { RenderHeadersAndFooters = renderStories, RenderHeaderFooterRegistry = registry }));
        Assert.DoesNotContain(html.Descendants(), e => (string?)e.Attribute("id") == "document-hf-registry");
        Assert.Equal(renderStories ? 1 : 0, html.Descendants(Xhtml.header).Count());
        Assert.Equal(renderStories ? 1 : 0, html.Descendants(Xhtml.footer).Count());
    }
}
