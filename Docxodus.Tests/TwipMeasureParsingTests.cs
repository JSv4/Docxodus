using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

public class TwipMeasureParsingTests
{
    [Theory]
    [InlineData("892.875", 892)]
    [InlineData("-892.875", -892)]
    [InlineData("2.54cm", 1440)]
    [InlineData("25.4mm", 1440)]
    [InlineData("3.2in", 4608)]
    [InlineData("31.25pt", 625)]
    [InlineData("2.5pc", 600)]
    [InlineData("2.5pi", 600)]
    [InlineData(" +892.875 ", 892)]
    [InlineData("2147483647.75", int.MaxValue)]
    [InlineData("-2147483648.75", int.MinValue)]
    public void StringsAndAttributesUseTheSameInvariantTwipConversion(string value, int expected)
    {
        var original = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
            Assert.Equal(expected, WordprocessingMLUtil.StringToTwips(value));
            Assert.Equal(expected, WordprocessingMLUtil.AttributeToTwips(new XAttribute(W.pos, value)));
        }
        finally { CultureInfo.CurrentCulture = original; }
    }

    [Theory]
    [InlineData("892,875")]
    [InlineData("31,25pt")]
    [InlineData("4px")]
    [InlineData("NaN")]
    [InlineData("Infinity")]
    [InlineData("2e3")]
    [InlineData("")]
    public void MalformedTwipsRemainInvalid(string value) =>
        Assert.Throws<FormatException>(() => WordprocessingMLUtil.StringToTwips(value));

    [Theory]
    [InlineData("2147483648")]
    [InlineData("-2147483649")]
    [InlineData("107374182.4pt")]
    public void IntegerConsumersKeepTheirRangeChecks(string value) =>
        Assert.Throws<OverflowException>(() => WordprocessingMLUtil.StringToTwips(value));

    private static WmlDocument AuthoredDocument(string tab, string rowHeight, string pageWidth, string pageHeight)
    {
        const string w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var styles = main.AddNewPart<StyleDefinitionsPart>();
            using var styleXml = new MemoryStream(Encoding.UTF8.GetBytes($"""
                <w:styles xmlns:w="{w}">
                  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
                  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
                  <w:style w:type="paragraph" w:styleId="MeasureTrack"><w:name w:val="Measure Track"/><w:basedOn w:val="Normal"/><w:pPr><w:tabs><w:tab w:val="left" w:pos="{tab}"/></w:tabs></w:pPr></w:style>
                </w:styles>
                """));
            styles.FeedData(styleXml);
            using var documentXml = new MemoryStream(Encoding.UTF8.GetBytes($"""
                <w:document xmlns:w="{w}"><w:body>
                  <w:p><w:pPr><w:pStyle w:val="MeasureTrack"/></w:pPr><w:r><w:t>Wren</w:t><w:tab/><w:t>field inventory.</w:t></w:r></w:p>
                  <w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
                    <w:tr><w:trPr><w:trHeight w:val="{rowHeight}" w:hRule="atLeast"/></w:trPr>
                      <w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Trail sample 07.</w:t></w:r></w:p></w:tc>
                      <w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>All items retained.</w:t></w:r></w:p></w:tc>
                    </w:tr>
                  </w:tbl>
                  <w:p><w:r><w:t>Field inspection complete.</w:t></w:r></w:p>
                  <w:sectPr><w:pgSz w:w="{pageWidth}" w:h="{pageHeight}"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
                </w:body></w:document>
                """));
            main.FeedData(documentXml);
        }
        return new WmlDocument("Wren.docx", stream.ToArray());
    }

    [Theory]
    [InlineData("1827.875", "34.3625pt")]
    [InlineData("91.39375pt", "687.25")]
    [InlineData("1827.875", "687.25")]
    public void FractionalStyleTabsAndRowHeightsRenderLikeTheirUnitControls(string tab, string height)
    {
        XElement Render(string tabValue, string rowValue, string width, string pageHeight) =>
            WmlToHtmlConverter.ConvertToHtml(AuthoredDocument(tabValue, rowValue, width, pageHeight),
                new WmlToHtmlConverterSettings { FabricateCssClasses = false });
        var actual = Render(tab, height, "12240.5", "15840.5");
        var expected = Render("91.39375pt", "34.3625pt", "612.025pt", "792.025pt");
        Assert.Equal(expected.Value, actual.Value);
        Assert.Equal(expected.Descendants().Attributes("style").Select(attribute => attribute.Value),
            actual.Descendants().Attributes("style").Select(attribute => attribute.Value));
        Assert.Contains("Trail sample 07.", actual.Value);
        Assert.Contains("All items retained.", actual.Value);
        Assert.Contains("Field inspection complete.", actual.Value);
    }
}
