using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Docxodus.Tests;

public class FootnoteSeparatorLineTests
{
    private static readonly XNamespace Wns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Theory]
    [InlineData(false, 11)]
    [InlineData(true, 11)]
    [InlineData(false, 17)]
    [InlineData(true, 17)]
    public void DefaultSeparatorReservesOneNaturalLineInItsOwnFont(bool explicitDefinitions, int fontSize)
    {
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
        var main = document.AddMainDocumentPart();
        main.PutXDocument(new XDocument(new XElement(Wns + "document", new XElement(Wns + "body",
            new XElement(Wns + "p", new XElement(Wns + "r", new XElement(Wns + "t", "Body")),
                new XElement(Wns + "r", new XElement(Wns + "footnoteReference", new XAttribute(Wns + "id", "1"))))))));
        main.AddNewPart<StyleDefinitionsPart>().PutXDocument(new XDocument(new XElement(Wns + "styles",
            new XElement(Wns + "docDefaults",
                new XElement(Wns + "rPrDefault", new XElement(Wns + "rPr",
                    new XElement(Wns + "rFonts", new XAttribute(Wns + "ascii", "Times New Roman"),
                        new XAttribute(Wns + "hAnsi", "Times New Roman")),
                    new XElement(Wns + "sz", new XAttribute(Wns + "val", fontSize * 2))))),
            new XElement(Wns + "style", new XAttribute(Wns + "type", "paragraph"),
                new XAttribute(Wns + "default", "1"), new XAttribute(Wns + "styleId", "Normal"),
                new XElement(Wns + "name", new XAttribute(Wns + "val", "Normal"))))));
        var notes = new XElement(Wns + "footnotes",
            new XElement(Wns + "footnote", new XAttribute(Wns + "id", "1"),
                new XElement(Wns + "p", new XElement(Wns + "r",
                    new XElement(Wns + "rPr", new XElement(Wns + "sz", new XAttribute(Wns + "val", "20"))),
                    new XElement(Wns + "t", "Note has a different font size")))));
        if (explicitDefinitions)
        {
            foreach (var (kind, id, mark) in new[] { ("separator", -1, "separator"), ("continuationSeparator", 0, "continuationSeparator") })
                notes.Add(new XElement(Wns + "footnote", new XAttribute(Wns + "type", kind), new XAttribute(Wns + "id", id),
                    new XElement(Wns + "p", new XElement(Wns + "pPr", new XElement(Wns + "spacing",
                        new XAttribute(Wns + "line", "1"), new XAttribute(Wns + "lineRule", "exact"))),
                        new XElement(Wns + "r", new XElement(Wns + mark)))));
        }
        main.AddNewPart<FootnotesPart>().PutXDocument(new XDocument(notes));
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(document));

        var html = WmlToHtmlConverter.ConvertToHtml(document, new WmlToHtmlConverterSettings
        {
            RenderFootnotesAndEndnotes = true,
            RenderPagination = PaginationMode.Paginated,
            FabricateCssClasses = false,
            StampAnchors = true,
            StampCanonicalSourceAnchors = true,
        });
        var registry = html.Descendants().Single(e => (string?)e.Attribute("id") == "pagination-footnote-registry");
        var separators = registry.Elements().Where(e => e.Attribute("data-footnote-separator") != null).ToArray();
        Assert.Equal(2, separators.Length);
        Assert.Single(registry.Elements(), e => e.Attribute("data-footnote-id") != null);
        foreach (var story in separators)
        {
            Assert.Contains("footnote-separator-line", (string?)story.Attribute("class") ?? "");
            var style = ((string?)story.Attribute("style") ?? "").Replace(" ", "", StringComparison.Ordinal);
            Assert.Contains("font-size:" + fontSize.ToString(CultureInfo.InvariantCulture) + "pt", style);
            Assert.Contains("TimesNewRoman", style);
            Assert.Contains("line-height:normal", style);
            Assert.Contains("height:1lh", style);
            Assert.Single(story.Elements());
            Assert.Equal("hr", story.Elements().Single().Name.LocalName);
            Assert.DoesNotContain(story.DescendantsAndSelf().Attributes(), a =>
                a.Name.LocalName is "id" or "data-anchor" or "data-source-anchor-id" or "data-footnote-id");
        }
    }
}
