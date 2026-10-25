using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

// Issue #1201: native Word leaves a measurable one-point sliver in a zero-width fixed cell.
public class ZeroWidthColumnTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // Literal original ledger; no external document or template is read.
    private static XElement Render(string narrowGridColumn, Action<XElement>? configure = null)
    {
        var cell = (string text, string width) =>
            $"<w:tc><w:tcPr><w:tcW w:w=\"{width}\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>";
        var content = $"<w:document xmlns:w=\"{W}\"><w:body><w:tbl><w:tblPr><w:tblW w:w=\"5400\" w:type=\"dxa\"/>"
            + "<w:tblLayout w:type=\"fixed\"/></w:tblPr><w:tblGrid>"
            + $"<w:gridCol {narrowGridColumn}/><w:gridCol w:w=\"5400\"/></w:tblGrid>"
            + $"<w:tr>{cell("Alder", "0")}{cell("Birch", "5400")}</w:tr></w:tbl>"
            + "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\"/></w:sectPr></w:body></w:document>";
        var root = XElement.Parse(content);
        configure?.Invoke(root.Descendants((XNamespace)W + "tbl").Single());
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new DocumentFormat.OpenXml.Wordprocessing.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings();
            using var xml = new MemoryStream(Encoding.UTF8.GetBytes(root.ToString()));
            main.FeedData(xml);
            PackageValidation.AssertValid(document, DocumentFormat.OpenXml.FileFormatVersions.Microsoft365);
        }
        return WmlToHtmlConverter.ConvertToHtml(new WmlDocument("Alder.docx", stream.ToArray()),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false });
    }

    private static string CellStyle(XElement html, string text) =>
        (string?)html.Descendants(Xhtml.td).Single(td => td.Value == text).Attribute("style") ?? "";

    [Fact]
    public void ZeroWidthGridColumnClipsItsCellHorizontally()
    {
        var html = Render("w:w=\"0\"");
        Assert.Contains("overflow-x: clip", CellStyle(html, "Alder"));
        Assert.Equal("1pt", Style(html.Descendants(Xhtml.td).First())["width"]);
        Assert.DoesNotContain("overflow-x", CellStyle(html, "Birch"));
    }

    [Theory]
    [InlineData("w:w=\"480\"")]
    [InlineData("")] // a grid column with no stated width is not a zero-width column
    public void OtherGridColumnsDoNotClip(string narrowGridColumn) =>
        Assert.DoesNotContain("overflow-x", CellStyle(Render(narrowGridColumn), "Alder"));

    private static Dictionary<string, string> Style(XElement element) =>
        ((string?)element.Attribute("style") ?? "").Split(';', StringSplitOptions.RemoveEmptyEntries)
            .Select(s => s.Split(':', 2)).ToDictionary(s => s[0].Trim(), s => s[1].Trim());

    [Theory]
    [InlineData("autofit", "0")]
    [InlineData("fixed", "480")]
    public void AutoFitAndPositiveCellWidthsAreNotForcedToOnePoint(string layout, string width)
    {
        XNamespace w = W;
        var html = Render("w:w=\"0\"", tbl => {
            tbl.Element(w + "tblPr")!.Element(w + "tblLayout")!.SetAttributeValue(w + "type", layout);
            tbl.Descendants(w + "tcW").First().SetAttributeValue(w + "w", width);
        });
        Assert.DoesNotContain("overflow-x", CellStyle(html, "Alder"));
        Assert.Equal(width == "480" ? "24pt" : "0pt", Style(html.Descendants(Xhtml.td).First())["width"]);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ASpanAcrossZeroColumnsHasOnePointTotalAndKeepsPadding(bool withPadding)
    {
        XNamespace w = W;
        var html = Render("w:w=\"0\"", tbl => {
            tbl.Element(w + "tblGrid")!.AddFirst(new XElement(w + "gridCol", new XAttribute(w + "w", 0)));
            tbl.Descendants(w + "tcPr").First().Add(new XElement(w + "gridSpan", new XAttribute(w + "val", 2)));
            if (withPadding) tbl.Descendants(w + "tcPr").First().Add(new XElement(w + "tcMar",
                new[] { "left", "right" }.Select(side => new XElement(w + side,
                    new XAttribute(w + "w", 240), new XAttribute(w + "type", "dxa")))));
        });
        var style = Style(html.Descendants(Xhtml.td).First());
        Assert.Equal("1pt", style["width"]);
        if (withPadding) { Assert.Equal("12pt", style["padding-left"]); Assert.Equal("12pt", style["padding-right"]); }
    }

    [Fact]
    public void GridBeforeAndCellsInsideControlsUseTheirOwnRowAndSkipNestedCells()
    {
        XNamespace w = W;
        var html = Render("w:w=\"0\"", tbl => {
            var row = tbl.Element(w + "tr")!;
            row.AddFirst(new XElement(w + "trPr", new XElement(w + "gridBefore", new XAttribute(w + "val", 1))));
            var prior = XElement.Parse($"<w:tc xmlns:w='{W}'><w:tcPr><w:tcW w:w='480' w:type='dxa'/></w:tcPr><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w='480'/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Nested</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>Cedar</w:t></w:r></w:p></w:tc>");
            row.Element(w + "trPr")!.AddAfterSelf(new XElement(w + "sdt", new XElement(w + "sdtContent", prior)));
            tbl.Element(w + "tblGrid")!.AddFirst(new XElement(w + "gridCol", new XAttribute(w + "w", 480)));
            tbl.Element(w + "tblGrid")!.AddFirst(new XElement(w + "gridCol", new XAttribute(w + "w", 240)));
        });
        Assert.Equal("1pt", Style(html.Descendants(Xhtml.td).Single(td => td.Value == "Alder"))["width"]);
        Assert.DoesNotContain("overflow-x", CellStyle(html, "Birch"));
        Assert.DoesNotContain("overflow-x", CellStyle(html, "Nested"));
    }

    [Theory]
    [InlineData("zero-exact", 1)]
    [InlineData("zero-atleast", 1)]
    [InlineData("width-control", 0)]
    [InlineData("tables-zero-width-column-times-new-roman", 3)]
    [InlineData("tables-zero-width-column-control-times-new-roman", 0)]
    [InlineData("zero-default-margins", 3)]
    [InlineData("zero-explicit-margins", 3)]
    [InlineData("zero-spanned-columns", 3)]
    [InlineData("zero-separate-columns", 6)]
    public void CapturedWordInputsKeepEveryZeroCellMeasurable(string id, int count)
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../TestFiles/ZeroWidthColumns"));
        var bytes = File.ReadAllBytes(Path.Combine(directory, id + ".docx"));
        using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "word-observations.json")));
        var observed = manifest.RootElement.GetProperty("cases").EnumerateArray().Single(c => c.GetProperty("id").GetString() == id);
        Assert.Equal(observed.GetProperty("inputSha256").GetString(), Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant());
        PackageValidation.AssertValid(bytes, DocumentFormat.OpenXml.FileFormatVersions.Microsoft365);
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument(id + ".docx", bytes), new WmlToHtmlConverterSettings { FabricateCssClasses = false });
        var cells = html.Descendants(Xhtml.td).Select(Style).Where(s => s.ContainsKey("overflow-x")).ToList();
        Assert.Equal(count, cells.Count);
        Assert.All(cells, s => { Assert.Equal("1pt", s["width"]); Assert.Equal("clip", s["overflow-x"]); });
    }
}
