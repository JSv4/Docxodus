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

    [Theory]
    [InlineData("4", 4)]
    [InlineData("4.5", 4)]
    [InlineData("4.9", 4)]
    [InlineData(" +12.75 ", 12)]
    [InlineData("4.", 4)]
    [InlineData("-0", 0)]
    public void BorderSizesUseTheInvariantMeasureSyntaxWithoutUnits(string value, int expected)
    {
        var original = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
            Assert.Equal(expected, WordprocessingMLUtil.BorderSize(new XElement(W.top, new XAttribute(W.sz, value))));
        }
        finally { CultureInfo.CurrentCulture = original; }
    }

    [Fact]
    public void AnOmittedBorderSizeIsZero()
    {
        Assert.Equal(0, WordprocessingMLUtil.BorderSize(new XElement(W.top, new XAttribute(W.val, "single"))));
        Assert.Equal(0, WordprocessingMLUtil.BorderSize(null));
    }

    [Theory]
    [InlineData("4,5")]
    [InlineData("4pt")]
    [InlineData("NaN")]
    [InlineData("Infinity")]
    [InlineData("1e1")]
    [InlineData("")]
    public void MalformedBorderSizesRemainInvalid(string value) =>
        Assert.Throws<FormatException>(() => WordprocessingMLUtil.BorderSize(new XElement(W.top, new XAttribute(W.sz, value))));

    [Fact]
    public void BorderSizesKeepTheIntegerRangeCheck() =>
        Assert.Throws<OverflowException>(() => WordprocessingMLUtil.BorderSize(new XElement(W.top, new XAttribute(W.sz, "2147483648"))));

    private static WmlDocument AuthoredDocument(string tab, string rowHeight, string pageWidth, string pageHeight) =>
        Document($"""
            <w:p><w:pPr><w:pStyle w:val="MeasureTrack"/></w:pPr><w:r><w:t>Wren</w:t><w:tab/><w:t>field inventory.</w:t></w:r></w:p>
            <w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
              <w:tr><w:trPr><w:trHeight w:val="{rowHeight}" w:hRule="atLeast"/></w:trPr>
                <w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Trail sample 07.</w:t></w:r></w:p></w:tc>
                <w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>All items retained.</w:t></w:r></w:p></w:tc>
              </w:tr>
            </w:tbl>
            <w:p><w:r><w:t>Field inspection complete.</w:t></w:r></w:p>
            <w:sectPr><w:pgSz w:w="{pageWidth}" w:h="{pageHeight}"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
            """,
            $"""<w:style w:type="paragraph" w:styleId="MeasureTrack"><w:name w:val="Measure Track"/><w:basedOn w:val="Normal"/><w:pPr><w:tabs><w:tab w:val="left" w:pos="{tab}"/></w:tabs></w:pPr></w:style>""");

    private static WmlDocument Document(string body, string extraStyles = "")
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
                  {extraStyles}
                </w:styles>
                """));
            styles.FeedData(styleXml);
            using var documentXml = new MemoryStream(Encoding.UTF8.GetBytes($"""<w:document xmlns:w="{w}"><w:body>{body}</w:body></w:document>"""));
            main.FeedData(documentXml);
        }
        return new WmlDocument("Wren.docx", stream.ToArray());
    }

    private static string Border(string side, string size) => $"""<w:{side} w:val="single" w:sz="{size}" w:space="0" w:color="000000"/>""";

    private static string Borders(string size) =>
        Border("top", size) + Border("left", size) + Border("bottom", size) + Border("right", size);

    // A minimum row height adds the cells' declared horizontal borders to the row.
    private static string MeasureTable(string tableBorders = "", string cellBorders = "", string firstColumn = "4000") => $"""
        <w:tbl><w:tblPr>{tableBorders}<w:tblLayout w:type="fixed"/></w:tblPr>
          <w:tblGrid><w:gridCol w:w="{firstColumn}"/><w:gridCol w:w="3000"/></w:tblGrid>
          <w:tr><w:trPr><w:trHeight w:val="600" w:hRule="atLeast"/></w:trPr><w:tc><w:tcPr>{cellBorders}</w:tcPr><w:p><w:r><w:t>Heron tally 12.</w:t></w:r></w:p></w:tc>
            <w:tc><w:tcPr>{cellBorders}</w:tcPr><w:p><w:r><w:t>Reed count kept.</w:t></w:r></w:p></w:tc></w:tr>
        </w:tbl>
        <w:p><w:r><w:t>Survey closed.</w:t></w:r></w:p>
        """;

    // Each carrier failed the whole conversion on a bare fractional value (issue #1622): borders
    // in every version, gridCol from 15.1.0 and section column spacing alike.
    public static TheoryData<string, string, string> FractionalWholeNumberMeasures => new()
    {
        { "paragraph border", Paragraph("4.5"), Paragraph("4") },
        { "table border", MeasureTable(tableBorders: TableBorders("4.5")), MeasureTable(tableBorders: TableBorders("4")) },
        { "cell border", MeasureTable(cellBorders: $"<w:tcBorders>{Borders("4.5")}</w:tcBorders>"),
            MeasureTable(cellBorders: $"<w:tcBorders>{Borders("4")}</w:tcBorders>") },
        { "grid column", MeasureTable(firstColumn: "4000.5"), MeasureTable(firstColumn: "4000") },
        { "column spacing", Columns("720.5"), Columns("720") },
    };

    private static string Paragraph(string size) =>
        $"""<w:p><w:pPr><w:pBdr>{Borders(size)}</w:pBdr></w:pPr><w:r><w:t>Heron ledger entry.</w:t></w:r></w:p>""";

    private static string TableBorders(string size) =>
        $"<w:tblBorders>{Borders(size)}{Border("insideH", size)}{Border("insideV", size)}</w:tblBorders>";

    private static string Columns(string space) => $"""
        <w:p><w:r><w:t>Heron column one.</w:t></w:r></w:p>
        <w:sectPr><w:cols w:num="2" w:space="{space}"/></w:sectPr>
        """;

    [Theory]
    [MemberData(nameof(FractionalWholeNumberMeasures))]
    public void FractionalWholeNumberMeasuresRenderLikeTheWholeValueWordReads(string carrier, string fractional, string whole)
    {
        static string Render(string body) => WmlToHtmlConverter.ConvertToHtml(Document(body),
            new WmlToHtmlConverterSettings { FabricateCssClasses = true }).ToString(SaveOptions.DisableFormatting);
        _ = carrier; // names the case in test output
        Assert.Equal(Render(whole), Render(fractional));
    }

    [Theory]
    [InlineData("NaN")]
    [InlineData("1e1")]
    [InlineData("4,5")]
    public void MalformedBorderSizesStillFailConversion(string size) =>
        Assert.Throws<FormatException>(() => WmlToHtmlConverter.ConvertToHtml(Document(Paragraph(size)),
            new WmlToHtmlConverterSettings()));

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
