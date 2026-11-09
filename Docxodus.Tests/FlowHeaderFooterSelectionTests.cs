using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Wp = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace Docxodus.Tests;

public class FlowHeaderFooterSelectionTests
{
    [Fact]
    public void PaginatedRegistryRetainsTheEmptyFinalSectionsReferences()
    {
        var html = XDocument.Parse(HtmlConversionOps.ConvertToHtml(Document("empty-final", false, false),
            new HtmlConversionOptions { RenderHeadersAndFooters = true, PaginationMode = 1 }));
        var headers = html.Descendants().Where(e => (string?)e.Attribute("data-hf-type") == "header-default").ToList();
        Assert.Equal(new[] { "0", "1", "2" }, headers.Select(e => (string?)e.Attribute("data-section")));
        Assert.Equal(new[] { "Old header", "Current header", "Old header" }, headers.Select(e => e.Value));
    }

    [Theory]
    [InlineData("none")]
    [InlineData("block")]
    [InlineData("row")]
    [InlineData("cell")]
    public void PaginatedRegistryKeepsSectionsInsideTransparentCustomXml(string wrapper)
    {
        var html = XDocument.Parse(HtmlConversionOps.ConvertToHtml(Document("inherited", false, false, wrapper),
            new HtmlConversionOptions { RenderHeadersAndFooters = true, PaginationMode = 1 }));
        var headers = html.Descendants().Where(e => (string?)e.Attribute("data-hf-type") == "header-default").ToList();
        Assert.Equal(new[] { "Old header", "Current header", "Current header" }, headers.Select(e => e.Value));
        Assert.Equal(new[] { "0", "1", "2" }, headers.Select(e => (string?)e.Attribute("data-section")));
        Assert.Equal(headers.Select(e => (string?)e.Attribute("data-section")),
            html.Descendants().Attributes("data-section-index").Select(a => a.Value));
    }

    private static byte[] Document(string scenario, bool reverseParts, bool enabled, string wrapper = "none")
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings(new Wp.EvenAndOddHeaders { Val = enabled });
            var parts = new[] { "Old", "Current", "Unused" };
            foreach (var name in reverseParts ? parts.Reverse() : parts)
            {
                main.AddNewPart<HeaderPart>("header" + name).Header =
                    new Wp.Header(new Wp.Paragraph(new Wp.Run(new Wp.Text(name + " header"))));
                main.AddNewPart<FooterPart>("footer" + name).Footer =
                    new Wp.Footer(new Wp.Paragraph(new Wp.Run(new Wp.Text(name + " footer"))));
            }
            Wp.SectionProperties Section(string name, Wp.HeaderFooterValues type) => new(
                new Wp.HeaderReference { Id = "header" + name, Type = type },
                new Wp.FooterReference { Id = "footer" + name, Type = type });
            var body = new Wp.Body();
            if (scenario is "inherited" or "empty-final")
            {
                body.Append(new Wp.Paragraph(new Wp.ParagraphProperties(Section("Old", Wp.HeaderFooterValues.Default)),
                    new Wp.Run(new Wp.Text("Opening"))));
                body.Append(new Wp.Paragraph(new Wp.ParagraphProperties(Section("Current", Wp.HeaderFooterValues.Default)),
                    new Wp.Run(new Wp.Text("Middle"))));
            }
            if (scenario != "empty-final")
                body.Append(new Wp.Paragraph(new Wp.Run(new Wp.Text("Body"))));
            var last = scenario switch
            {
                "explicit" => Section("Current", Wp.HeaderFooterValues.Default),
                "empty-final" => Section("Old", Wp.HeaderFooterValues.Default),
                "first" or "inherited" => Section("Unused", Wp.HeaderFooterValues.First),
                "even" => Section("Unused", Wp.HeaderFooterValues.Even),
                _ => new Wp.SectionProperties(),
            };
            last.Append(new Wp.TitlePage { Val = enabled });
            body.Append(last);
            if (wrapper != "none")
            {
                var first = body.GetFirstChild<Wp.Paragraph>()!;
                first.Remove();
                OpenXmlElement container = wrapper switch
                {
                    "row" => new Wp.Table(new Wp.CustomXmlRow(new Wp.TableRow(new Wp.TableCell(first))) { Element = "Section" }),
                    "cell" => new Wp.Table(new Wp.TableRow(new Wp.CustomXmlCell(new Wp.TableCell(first)) { Element = "Section" })),
                    _ => new Wp.CustomXmlBlock(first) { Element = "Section" },
                };
                body.PrependChild(container);
            }
            main.Document = new Wp.Document(body);
        }
        return stream.ToArray();
    }

    [Theory]
    [InlineData("inherited", false, false)]
    [InlineData("inherited", true, true)]
    [InlineData("explicit", false, false)]
    [InlineData("explicit", true, true)]
    [InlineData("first", false, false)]
    [InlineData("first", true, true)]
    [InlineData("even", false, false)]
    [InlineData("even", true, true)]
    [InlineData("unreferenced", false, false)]
    [InlineData("unreferenced", true, true)]
    [InlineData("empty-final", false, false)]
    [InlineData("empty-final", true, true)]
    public void FlowUsesOnlyTheFinalSectionsEffectiveDefault(string scenario, bool reverseParts, bool enabled)
    {
        var html = XDocument.Parse(HtmlConversionOps.ConvertToHtml(Document(scenario, reverseParts, enabled),
            new HtmlConversionOptions { RenderHeadersAndFooters = true }));
        foreach (var kind in new[] { "header", "footer" })
        {
            var stories = html.Descendants(Xhtml.xhtml + kind).ToList();
            if (scenario is "inherited" or "explicit" or "empty-final")
                Assert.Equal((scenario == "empty-final" ? "Old " : "Current ") + kind, Assert.Single(stories).Value);
            else
                Assert.Empty(stories);
        }
        Assert.Contains(scenario == "empty-final" ? "Middle" : "Body", html.Root!.Value);
    }
}
