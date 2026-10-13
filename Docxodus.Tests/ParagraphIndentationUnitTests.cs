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

public class ParagraphIndentationUnitTests
{
    private static byte[] AuthoredDocument(string value)
    {
        const string w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        string Paragraph(string marker, string attributes, bool border = false, bool bidi = false) =>
            "<w:p><w:pPr>" + (border ? "<w:pBdr><w:bottom w:val=\"single\" w:sz=\"8\"/></w:pBdr>" : "")
            + (bidi ? "<w:bidi/>" : "") + $"<w:ind {attributes}/></w:pPr><w:r><w:t>{marker}</w:t></w:r></w:p>";
        var body = Paragraph("Cedar physical edges.", $"w:left=\"{value}\" w:right=\"{value}\"")
            + Paragraph("Cedar logical edges.", $"w:start=\"{value}\" w:end=\"{value}\" w:left=\"1440\" w:right=\"1440\"", bidi: true)
            + Paragraph("Cedar first line.", $"w:left=\"{value}\" w:firstLine=\"{value}\"")
            + Paragraph("Cedar bordered hanging line.", $"w:left=\"{value}\" w:hanging=\"{value}\"", border: true);
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            using var xml = new MemoryStream(Encoding.UTF8.GetBytes($"<w:document xmlns:w=\"{w}\"><w:body>{body}<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:left=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:header=\"720\" w:footer=\"720\"/></w:sectPr></w:body></w:document>"));
            main.FeedData(xml);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc));
        }
        return stream.ToArray();
    }

    [Theory]
    [InlineData("36pt")]
    [InlineData("0.5in")]
    [InlineData("1.27cm")]
    [InlineData("12.7mm")]
    [InlineData("3pc")]
    [InlineData("3pi")]
    public void UniversalIndentationUnitsRenderLikeEquivalentTwips(string value)
    {
        XElement Render(string v) => WmlToHtmlConverter.ConvertToHtml(new WmlDocument("Cedar.docx", AuthoredDocument(v)),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false });
        var actual = Render(value);
        var expected = Render("720");
        Assert.Equal(expected.Value, actual.Value);
        Assert.Equal(expected.Descendants().Attributes("style").Select(a => a.Value),
            actual.Descendants().Attributes("style").Select(a => a.Value));
    }

    [Theory]
    [InlineData("36pt", 720)]
    [InlineData("0.5in", 720)]
    [InlineData("1.27cm", 720)]
    [InlineData("12.7mm", 720)]
    [InlineData("3pc", 720)]
    [InlineData("3pi", 720)]
    [InlineData("-12.7mm", -720)]
    [InlineData("-1.27cm", -720)]
    [InlineData("-0.5in", -720)]
    [InlineData("-36pt", -720)]
    [InlineData("-3pc", -720)]
    [InlineData("-3pi", -720)]
    [InlineData("12.996749877929688", 12)]
    [InlineData("-12.996749877929688", -12)]
    public void TwipConversionUsesInvariantUnitsAndPreservesIntegerTruncation(string value, int expected)
    {
        var original = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
            Assert.Equal(expected, WordprocessingMLUtil.AttributeToTwips(new XAttribute("measure", value)));
        }
        finally { CultureInfo.CurrentCulture = original; }
    }

    [Theory]
    [InlineData("12px")]
    [InlineData("12PT")]
    [InlineData("12,5pt")]
    [InlineData("1e2pt")]
    [InlineData("NaN")]
    [InlineData("")]
    [InlineData("pt")]
    public void InvalidMeasurementsStillFail(string value) => Assert.Throws<FormatException>(() =>
        WordprocessingMLUtil.AttributeToTwips(new XAttribute("measure", value)));

    [Fact]
    public void MissingMeasurementRemainsNull() => Assert.Null(WordprocessingMLUtil.AttributeToTwips(null));

    [Fact]
    public void CssMeasurementsRetainTheExistingFractionalTwipPrecision()
    {
        var measure = new XAttribute("measure", "712.7");
        Assert.Equal(712.7m, WordprocessingMLUtil.AttributeToDecimalTwips(measure));
        Assert.Equal(712, WordprocessingMLUtil.AttributeToTwips(measure));
    }
}
