#nullable enable
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Docxodus.Tests;

public class TableRowHeightUnitTests
{
    private static byte[] AuthoredDocument(string? value, string rule = "atLeast")
    {
        const string w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var height = $"<w:trHeight {(value == null ? "" : $"w:val=\"{value}\"")} w:hRule=\"{rule}\"/>";
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            using var xml = new MemoryStream(Encoding.UTF8.GetBytes($"<w:document xmlns:w=\"{w}\"><w:body><w:tbl><w:tblPr><w:tblW w:w=\"7200\" w:type=\"dxa\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"7200\"/></w:tblGrid><w:tr><w:trPr>{height}</w:trPr><w:tc><w:tcPr><w:tcW w:w=\"7200\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>Alder row record.</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/></w:sectPr></w:body></w:document>"));
            main.FeedData(xml);
            // The SDK models trHeight/@val as UInt32 and rejects the schema's unit branch.
            // Independently validate the whole-twip control here; full ECMA XSD evidence
            // for the unit packages accompanies the product and benchmark PRs.
            if (value == "720" || value == null)
                Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc));
        }
        return stream.ToArray();
    }

    private static XElement Render(string? value, string rule = "atLeast") =>
        WmlToHtmlConverter.ConvertToHtml(new WmlDocument("Alder.docx", AuthoredDocument(value, rule)),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false });

    [Theory]
    [InlineData("36pt", "atLeast")]
    [InlineData("0.5in", "atLeast")]
    [InlineData("1.27cm", "atLeast")]
    [InlineData("12.7mm", "atLeast")]
    [InlineData("3pc", "atLeast")]
    [InlineData("3pi", "atLeast")]
    [InlineData("36pt", "exact")]
    public void ValidUnitsRenderLikeWholeTwips(string value, string rule)
    {
        var actual = Render(value, rule);
        var expected = Render("720", rule);
        Assert.Equal(expected.Value, actual.Value);
        Assert.Equal(expected.Descendants().Attributes("style").Select(a => a.Value),
            actual.Descendants().Attributes("style").Select(a => a.Value));
    }

    // Issue #1121: row heights keep twip precision. The former 0.01in quantum made 55pt rows
    // 54.72pt and mapped both 25.075pt and 25.1pt to 25.2pt.
    [Theory]
    [InlineData("25.075pt", "height: 25.075pt")]
    [InlineData("502", "height: 25.1pt")]
    [InlineData("1100", "height: 55pt")]
    [InlineData("720", "height: 36pt")]
    [InlineData("687.25", "height: 34.3625pt")]
    public void RowHeightsKeepTwipPrecision(string value, string expected) =>
        Assert.Contains(expected, (string?)Render(value).Descendants(Xhtml.tr).Single().Attribute("style"));

    [Theory]
    [InlineData("atLeast")]
    [InlineData("exact")]
    [InlineData("auto")]
    public void FractionalTwipsKeepTheHeightRule(string rule)
    {
        var actual = Render("687.25", rule);
        var expected = Render("34.3625pt", rule);
        Assert.Equal(expected.Value, actual.Value);
        Assert.Equal(expected.Descendants().Attributes("style").Select(attribute => attribute.Value),
            actual.Descendants().Attributes("style").Select(attribute => attribute.Value));
    }

    [Theory]
    [InlineData("687,25")]
    [InlineData("NaN")]
    [InlineData("Infinity")]
    [InlineData("2147483647.25")]
    [InlineData("-1")]
    [InlineData("-1.25")]
    [InlineData("-1pt")]
    [InlineData("12px")]
    [InlineData("12PT")]
    [InlineData("12,5pt")]
    [InlineData("1e2pt")]
    [InlineData("+36pt")]
    [InlineData(" 36pt ")]
    [InlineData("2147483648")]
    [InlineData("9999999999999999999999999999999999999999pt")]
    [InlineData("")]
    public void InvalidOrUnrepresentableHeightNamesTheProperty(string value)
    {
        var error = Assert.Throws<FormatException>(() => Render(value));
        Assert.Contains("w:trHeight/@w:val", error.Message);
        Assert.Contains("non-negative", error.Message);
    }

    [Fact]
    public void MissingHeightDoesNotAddCssHeight() =>
        Assert.DoesNotContain("height:", (string?)Render(null).Descendants(Xhtml.tr).Single().Attribute("style") ?? "");

    [Theory]
    [InlineData("+720")]
    [InlineData(" 720 ")]
    public void WholeTwipLexicalFormsRemainEquivalent(string value) =>
        Assert.Equal((string?)Render("720").Descendants(Xhtml.tr).Single().Attribute("style"),
            (string?)Render(value).Descendants(Xhtml.tr).Single().Attribute("style"));

    [Fact]
    public void UnitHeightUsesInvariantCulture()
    {
        var original = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
            Assert.Equal((string?)Render("720").Descendants(Xhtml.tr).Single().Attribute("style"),
                (string?)Render("12.7mm").Descendants(Xhtml.tr).Single().Attribute("style"));
            Assert.Equal((string?)Render("34.3625pt").Descendants(Xhtml.tr).Single().Attribute("style"),
                (string?)Render("687.25").Descendants(Xhtml.tr).Single().Attribute("style"));
        }
        finally { CultureInfo.CurrentCulture = original; }
    }
}
