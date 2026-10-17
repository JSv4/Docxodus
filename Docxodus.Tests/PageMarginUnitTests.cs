#nullable enable
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

public class PageMarginUnitTests
{
    private static WmlDocument AuthoredDocument(string edge, string band, string? left = null,
        string? top = null, bool omitMargins = false)
    {
        const string w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var margins = omitMargins ? "" : $"<w:pgMar w:top=\"{top ?? edge}\" w:right=\"{edge}\" w:bottom=\"{edge}\" w:left=\"{left ?? edge}\" w:header=\"{band}\" w:footer=\"{band}\" w:gutter=\"0\"/>";
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            using var xml = new MemoryStream(Encoding.UTF8.GetBytes($"<w:document xmlns:w=\"{w}\"><w:body><w:p><w:r><w:t>Juniper margin inventory.</w:t></w:r></w:p><w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/>{margins}</w:sectPr></w:body></w:document>"));
            main.FeedData(xml);
        }
        return new WmlDocument("Juniper.docx", stream.ToArray());
    }

    private static double[] Dimensions(SectionMetadata section) => new[] {
        section.PageWidthPt, section.PageHeightPt, section.MarginTopPt, section.MarginRightPt,
        section.MarginBottomPt, section.MarginLeftPt, section.HeaderPt, section.FooterPt,
        section.ContentWidthPt, section.ContentHeightPt };

    [Theory]
    [InlineData("54pt", "18pt", PaginationMode.None)]
    [InlineData("0.75in", "0.25in", PaginationMode.None)]
    [InlineData("1.905cm", "0.635cm", PaginationMode.None)]
    [InlineData("19.05mm", "6.35mm", PaginationMode.None)]
    [InlineData("4.5pc", "1.5pc", PaginationMode.None)]
    [InlineData("4.5pi", "1.5pi", PaginationMode.None)]
    [InlineData("54pt", "18pt", PaginationMode.Paginated)]
    [InlineData("0.75in", "0.25in", PaginationMode.Paginated)]
    [InlineData("1.905cm", "0.635cm", PaginationMode.Paginated)]
    [InlineData("19.05mm", "6.35mm", PaginationMode.Paginated)]
    [InlineData("4.5pc", "1.5pc", PaginationMode.Paginated)]
    [InlineData("4.5pi", "1.5pi", PaginationMode.Paginated)]
    public void MarginUnitsAgreeInMetadataAndBothHtmlModes(string edge, string band, PaginationMode mode)
    {
        var actual = AuthoredDocument(edge, band);
        var control = AuthoredDocument("1080", "360");
        Assert.Equal(Dimensions(WmlToHtmlConverter.GetDocumentMetadata(control).Sections.Single()),
            Dimensions(WmlToHtmlConverter.GetDocumentMetadata(actual).Sections.Single()));
        var settings = new WmlToHtmlConverterSettings { RenderPagination = mode, GeneratePageCss = true };
        var expectedHtml = WmlToHtmlConverter.ConvertToHtml(control, settings);
        var actualHtml = WmlToHtmlConverter.ConvertToHtml(actual, settings);
        string[] Layout(XElement html) => html.Descendants().Attributes()
            .Where(a => a.Name.LocalName.StartsWith("data-margin-", StringComparison.Ordinal)
                || a.Name.LocalName.StartsWith("data-content-", StringComparison.Ordinal)
                || a.Name.LocalName is "data-header-height" or "data-footer-height")
            .Select(a => a.ToString()).ToArray();
        Assert.NotEmpty(Layout(actualHtml));
        Assert.Equal(Layout(expectedHtml), Layout(actualHtml));
        Assert.Equal(expectedHtml.Descendants(Xhtml.style).Select(e => e.Value),
            actualHtml.Descendants(Xhtml.style).Select(e => e.Value));
    }

    [Fact]
    public void MillimeterMarginRetainsPrecisionInHtmlGeometry()
    {
        var doc = AuthoredDocument("1440", "720", left: "25mm");
        var section = WmlToHtmlConverter.GetDocumentMetadata(doc).Sections.Single();
        Assert.Equal(25.0 * 72 / 25.4, section.MarginLeftPt, precision: 10);
        var html = WmlToHtmlConverter.ConvertToHtml(doc, new WmlToHtmlConverterSettings());
        // The geometry attribute carries the four decimals the page CSS does (issue #1164), not one.
        Assert.Contains(html.Descendants().Attributes("data-margin-left"), a => a.Value == "70.8661");
    }

    [Fact]
    public void MissingMarginsRetainDefaultPageGeometry() => Assert.Equal(
        new[] { 612d, 792d, 72d, 72d, 72d, 72d, 36d, 36d, 468d, 648d },
        Dimensions(WmlToHtmlConverter.GetDocumentMetadata(AuthoredDocument("", "", omitMargins: true)).Sections.Single()));

    [Theory]
    [InlineData("left")]
    [InlineData("top")]
    public void MalformedPresentMarginsAreExplicitFailures(string attribute)
    {
        var doc = AuthoredDocument("1440", "720", left: attribute == "left" ? "12px" : null,
            top: attribute == "top" ? "12px" : null);
        Assert.Throws<FormatException>(() => WmlToHtmlConverter.GetDocumentMetadata(doc));
        Assert.Throws<FormatException>(() => WmlToHtmlConverter.ConvertToHtml(doc, new WmlToHtmlConverterSettings()));
    }

    [Fact]
    public void SignedTopMarginsKeepTheirValue()
    {
        var section = WmlToHtmlConverter.GetDocumentMetadata(AuthoredDocument("1440", "720", top: "-18pt")).Sections.Single();
        Assert.Equal(-18d, section.MarginTopPt);
    }
}
