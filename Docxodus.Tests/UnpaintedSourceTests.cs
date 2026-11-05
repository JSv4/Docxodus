using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Xunit;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

namespace Docxodus.Tests;

public class UnpaintedSourceTests
{
    [Fact]
    public void RowlessTableKeepsItsIdentityWithoutRequiringPaintedGeometry()
    {
        var html = Convert(new Wp.Table(new Wp.TableProperties(), new Wp.TableGrid()),
            Paragraph("Comet observation."));
        var table = Assert.Single(html.Descendants(Xhtml.table));
        Assert.NotNull(table.Attribute("data-source-anchor-id"));
        Assert.Equal("true", (string?)table.Attribute("data-page-map-exclude"));
        Assert.Contains("Comet observation.", html.Value);
    }

    [Fact]
    public void EmptyCellParagraphStillRequiresItsVisibleLine()
    {
        var html = Convert(new Wp.Table(new Wp.TableProperties(),
            new Wp.TableGrid(new Wp.GridColumn { Width = "1800" }),
            new Wp.TableRow(new Wp.TableCell(new Wp.Paragraph()))));
        foreach (var source in html.Descendants().Where(e => e.Attribute("data-source-anchor-id") != null))
            Assert.Null(source.Attribute("data-page-map-exclude"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void JoinedHiddenParagraphHasNoPaintedFragmentButKeepsItsContent(bool emptyVisibleRun)
    {
        var hidden = Paragraph("Internal telescope note.");
        hidden.PrependChild(new Wp.ParagraphProperties(new Wp.ParagraphMarkRunProperties(new Wp.Vanish())));
        hidden.GetFirstChild<Wp.Run>()!.PrependChild(new Wp.RunProperties(new Wp.Vanish()));
        if (emptyVisibleRun) hidden.AppendChild(new Wp.Run(new Wp.Text("")));
        var first = Paragraph("Observation:");
        first.PrependChild(new Wp.ParagraphProperties(new Wp.ParagraphMarkRunProperties(new Wp.SpecVanish())));
        var html = Convert(first, hidden, Paragraph("Orbit verified."));

        var source = Assert.Single(html.Descendants(Xhtml.span), e =>
            e.Attribute("data-source-anchor-id") != null && e.Value.Contains("Internal telescope note."));
        Assert.Equal("true", (string?)source.Attribute("data-page-map-exclude"));
        Assert.Contains("Orbit verified.", html.Value);
    }

    [Fact]
    public void JoinedVisibleParagraphStillRequiresItsFragment()
    {
        var first = Paragraph("Observation:");
        first.PrependChild(new Wp.ParagraphProperties(new Wp.ParagraphMarkRunProperties(new Wp.SpecVanish())));
        var html = Convert(first, Paragraph("Orbit verified."));
        var source = Assert.Single(html.Descendants(Xhtml.span), e =>
            e.Attribute("data-source-anchor-id") != null && e.Value.Contains("Orbit verified."));
        Assert.Null(source.Attribute("data-page-map-exclude"));
    }

    [Fact]
    public void JoinedEmptyParagraphWithSpecialHiddenMarkKeepsItsPlaceholderFragment()
    {
        var first = Paragraph("Observation:");
        first.PrependChild(new Wp.ParagraphProperties(new Wp.ParagraphMarkRunProperties(new Wp.SpecVanish())));
        var empty = new Wp.Paragraph(new Wp.ParagraphProperties(
            new Wp.ParagraphMarkRunProperties(new Wp.SpecVanish())));
        var html = Convert(first, empty, Paragraph("Orbit verified."));
        foreach (var source in html.Descendants().Where(e => e.Attribute("data-source-anchor-id") != null))
            Assert.Null(source.Attribute("data-page-map-exclude"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StandaloneHiddenParagraphWithoutVisibleMarkDoesNotRequireGeometry(bool inCell)
    {
        var hidden = new Wp.Paragraph(new Wp.ParagraphProperties(
            new Wp.ParagraphMarkRunProperties(new Wp.Vanish())));
        var html = inCell
            ? Convert(new Wp.Table(new Wp.TableRow(new Wp.TableCell(Paragraph("Orbit verified."), hidden))))
            : Convert(Paragraph("Orbit verified."), hidden);
        var source = Assert.Single(html.Descendants(Xhtml.p), e => e.Attribute("data-source-anchor-id") != null && e.Value.Trim().Length == 0);
        Assert.NotNull(source.Attribute("data-source-anchor-id"));
        Assert.Equal("true", (string?)source.Attribute("data-page-map-exclude"));
        Assert.Contains("Orbit verified.", html.Value);
    }

    [Fact]
    public void HiddenFirstParagraphDoesNotExcludeJoinedVisibleContent()
    {
        var first = new Wp.Paragraph(new Wp.ParagraphProperties(
            new Wp.ParagraphMarkRunProperties(new Wp.Vanish())));
        var html = Convert(first, Paragraph("Orbit verified."));
        foreach (var source in html.Descendants().Where(e => e.Attribute("data-source-anchor-id") != null))
            Assert.Null(source.Attribute("data-page-map-exclude"));
    }

    [Fact]
    public void JoinedVisibleShapeWithHiddenMarkStillRequiresItsFragment()
    {
        var first = Paragraph("Observation:");
        first.PrependChild(new Wp.ParagraphProperties(new Wp.ParagraphMarkRunProperties(new Wp.SpecVanish())));
        var shape = new Wp.Paragraph(new Wp.ParagraphProperties(
            new Wp.ParagraphMarkRunProperties(new Wp.Vanish())), new Wp.Run(
                "<w:r xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:v=\"urn:schemas-microsoft-com:vml\">" +
                "<w:pict><v:rect style=\"width:24pt;height:12pt\" fillcolor=\"#219663\"/></w:pict></w:r>"));
        var html = Convert(first, shape, Paragraph("Orbit verified."));
        var source = Assert.Single(html.Descendants(Xhtml.span), e => e.Attribute("data-source-anchor-id") != null &&
            e.Descendants().Any(child => child.Name.LocalName == "svg"));
        Assert.Null(source.Attribute("data-page-map-exclude"));
    }

    [Fact]
    public void HiddenSectionBoundaryRetainsSectionButDoesNotRequirePaintedLine()
    {
        var hidden = new Wp.Paragraph(new Wp.ParagraphProperties(
            new Wp.ParagraphMarkRunProperties(new Wp.Vanish()),
            new Wp.SectionProperties(new Wp.PageSize { Width = 6000, Height = 7200 })));
        var html = Convert(Paragraph("Orbit verified."), hidden);
        var source = Assert.Single(html.Descendants(Xhtml.p), e =>
            e.Attribute("data-source-anchor-id") != null && e.Value.Trim().Length == 0);
        Assert.Equal("true", (string?)source.Attribute("data-page-map-exclude"));
        Assert.Contains("size: 300pt 360pt", html.Value);
    }

    private static Wp.Paragraph Paragraph(string text) => new(new Wp.Run(new Wp.Text(text)));

    private static XElement Convert(params OpenXmlElement[] blocks)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            main.Document = new Wp.Document(new Wp.Body(blocks));
        }
        return XElement.Parse(HtmlConversionOps.ConvertToHtml(stream.ToArray(), new HtmlConversionOptions
        {
            PaginationMode = (int)PaginationMode.Paginated,
            FabricateCssClasses = false,
        }));
    }
}
