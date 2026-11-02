using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

public class TableExactBorderTests
{
    private static string Cell(string text, int size, int span = 1, string properties = "", string attributes = "", string type = "single") =>
        $"<w:tc><w:tcPr><w:tcW w:w=\"{2400 * span}\" w:type=\"dxa\"/>"
        + (span > 1 ? $"<w:gridSpan w:val=\"{span}\"/>" : "") + properties + "<w:tcBorders>"
        + string.Concat(new[] { "top", "left", "bottom", "right" }.Select(side =>
            $"<w:{side} w:val=\"{type}\" w:sz=\"{size}\" {attributes}/>"))
        + $"</w:tcBorders></w:tcPr><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>";

    private static string Table(string rows) => "<w:tbl><w:tblPr><w:tblW w:w=\"4800\" w:type=\"dxa\"/>"
        + "<w:tblLayout w:type=\"fixed\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"2400\"/><w:gridCol w:w=\"2400\"/></w:tblGrid>"
        + rows + "</w:tbl>";

    private static XElement Convert(string table, bool header = false)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            Feed(main.AddNewPart<StyleDefinitionsPart>(), $"<w:styles xmlns:w=\"{W.w}\"/>");
            Feed(main.AddNewPart<DocumentSettingsPart>(), $"<w:settings xmlns:w=\"{W.w}\"/>");
            if (header) Feed(main.AddNewPart<HeaderPart>("header"), $"<w:hdr xmlns:w=\"{W.w}\">{table}<w:p/></w:hdr>");
            Feed(main, $"<w:document xmlns:w=\"{W.w}\" xmlns:r=\"{R.r}\"><w:body>{(header ? "<w:p/>" : table)}<w:sectPr>"
                + (header ? "<w:headerReference w:type=\"default\" r:id=\"header\"/>" : "")
                + "<w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:left=\"1440\" w:bottom=\"1440\" w:right=\"1440\" w:header=\"360\" w:footer=\"360\"/>"
                + "</w:sectPr></w:body></w:document>");
            PackageValidation.AssertValid(document, DocumentFormat.OpenXml.FileFormatVersions.Microsoft365);
        }
        return WmlToHtmlConverter.ConvertToHtml(new WmlDocument("Cypress.docx", stream.ToArray()),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false, RenderHeadersAndFooters = true });
    }

    private static void Feed(OpenXmlPart part, string xml)
    {
        var ordered = WordprocessingMLUtil.WmlOrderElementsPerStandard(XElement.Parse(xml));
        using var data = new MemoryStream(Encoding.UTF8.GetBytes(ordered.ToString()!));
        part.FeedData(data);
    }

    private static List<JsonElement> Edges(XElement row)
    {
        using var json = JsonDocument.Parse((string?)row.Attribute("data-docx-table-border-edges") ?? "[]");
        return json.RootElement.EnumerateArray().Select(e => e.Clone()).ToList();
    }

    [Theory]
    [InlineData(2, false)]
    [InlineData(4, false)]
    [InlineData(8, false)]
    [InlineData(12, false)]
    [InlineData(2, true)]
    public void Physical_widths_are_preserved_separately_from_the_native_CSS_fallback(int size, bool header)
    {
        var html = Convert(Table("<w:tr>" + Cell("CypressNorth", size) + Cell("CypressSouth", size) + "</w:tr>"), header);
        var table = Assert.Single(html.Descendants(Xhtml.table));
        Assert.NotNull(table.Attribute("data-docx-exact-table-borders"));
        var edges = Edges(Assert.Single(table.Descendants(Xhtml.tr)));
        Assert.Equal(7, edges.Count);
        Assert.All(edges, edge => Assert.Equal(size / 8m, edge[3].GetDecimal()));
        Assert.All(table.Descendants(Xhtml.td), cell => Assert.Contains("border-left: solid", (string?)cell.Attribute("style") ?? ""));
        Assert.Contains("CypressNorth", table.Value);
        Assert.Contains("CypressSouth", table.Value);
    }

    [Fact]
    public void A_colspan_keeps_distinct_winners_on_its_shared_bottom_segments()
    {
        var table = Table("<w:tr>" + Cell("CypressMerged", 2, 2) + "</w:tr><w:tr>"
            + Cell("CypressWest", 4) + Cell("CypressEast", 8) + "</w:tr>");
        var rows = Convert(table).Descendants(Xhtml.tr).ToList();
        Assert.Equal(new[] { .5m, 1m }, Edges(rows[0]).Where(e => e[0].GetString() == "bottom").Select(e => e[3].GetDecimal()));
        Assert.Equal(new[] { .5m, 1m }, Edges(rows[1]).Where(e => e[0].GetString() == "top").Select(e => e[3].GetDecimal()));
    }

    [Fact]
    public void A_vertical_merge_does_not_gain_an_internal_horizontal_edge()
    {
        var table = Table("<w:tr>" + Cell("CypressTall", 2, properties: "<w:vMerge w:val=\"restart\"/>")
            + Cell("CypressTop", 4) + "</w:tr><w:tr>" + Cell("", 2, properties: "<w:vMerge/>") + Cell("CypressBottom", 8) + "</w:tr>");
        var html = Convert(table);
        Assert.Single(html.Descendants(Xhtml.td), c => (string?)c.Attribute("rowspan") == "2");
        var rows = html.Descendants(Xhtml.tr).ToList();
        Assert.All(Edges(rows[0]).Where(e => e[0].GetString() == "bottom"), edge => Assert.Equal(1, edge[1].GetInt32()));
        Assert.All(Edges(rows[1]).Where(e => e[0].GetString() == "top"), edge => Assert.Equal(1, edge[1].GetInt32()));
    }

    [Fact]
    public void Winning_color_does_not_inherit_the_losing_borders_theme()
    {
        var table = Table("<w:tr>" + Cell("CypressTheme", 2, attributes: "w:themeColor=\"accent1\" w:color=\"FF0000\"")
            + Cell("CypressBlue", 8, attributes: "w:color=\"0000FF\"") + "</w:tr>");
        var edge = Assert.Single(Edges(Convert(table).Descendants(Xhtml.tr).Single()),
            e => e[0].GetString() == "vertical" && e[1].GetInt32() == 1);
        Assert.Equal("#0000FF", edge[4].GetString());
    }

    [Fact]
    public void Rtl_header_and_body_tables_resolve_the_same_physical_edges()
    {
        var cell = Cell("CypressRtl", 2, 2).Replace("<w:right w:val=\"single\" w:sz=\"2\"", "<w:right w:val=\"single\" w:sz=\"8\"");
        var table = Table("<w:tr>" + cell + "</w:tr>").Replace("<w:tblPr>", "<w:tblPr><w:bidiVisual/>");
        var body = Edges(Convert(table).Descendants(Xhtml.tr).Single()).Select(e => e.GetRawText());
        var header = Edges(Convert(table, true).Descendants(Xhtml.tr).Single()).Select(e => e.GetRawText());
        Assert.Equal(body, header);
    }

    [Fact]
    public void A_repeated_header_resolves_its_new_body_neighbor_from_original_borders()
    {
        var table = Table("<w:tr><w:trPr><w:tblHeader/></w:trPr>"
            + Cell("AlderHeader", 24, 2, attributes: "w:color=\"C02020\"") + "</w:tr><w:tr>"
            + Cell("AlderFirst", 2, 2, attributes: "w:color=\"205080\"") + "</w:tr><w:tr>"
            + Cell("AlderNext", 4, 2, attributes: "w:color=\"205080\"") + "</w:tr>");
        var rows = Convert(table).Descendants(Xhtml.tr).ToList();
        var original = Assert.Single(Edges(rows[2]), e => e[0].GetString() == "top");
        Assert.Equal(.5m, original[3].GetDecimal());
        Assert.Equal("#205080", original[4].GetString());
        using var json = JsonDocument.Parse((string)rows[2].Attribute("data-docx-table-header-edges")!);
        var continuation = Assert.Single(json.RootElement.EnumerateArray());
        Assert.Equal(3m, continuation[3].GetDecimal());
        Assert.Equal("#C02020", continuation[4].GetString());
    }

    [Fact]
    public void Tab_after_boundary_text_uses_the_same_content_width_as_bordered_cell_layout()
    {
        string Fixture(int gridWidth, bool border)
        {
            var cell = XElement.Parse(Cell("Cypress", 24).Replace("<w:tc>", $"<w:tc xmlns:w=\"{W.w}\">"));
            var properties = cell.Element(W.tcPr)!;
            if (!border) properties.Element(W.tcBorders)!.Remove();
            properties.Element(W.tcW)!.SetAttributeValue(W._w, gridWidth);
            properties.Add(new XElement(W.tcMar, new[] { W.left, W.right }.Select(side =>
                new XElement(side, new XAttribute(W._w, 0), new XAttribute(W.type, "dxa")))));
            cell.Element(W.p)!.ReplaceWith(XElement.Parse($"<w:p xmlns:w=\"{W.w}\"><w:r><w:rPr>"
                + "<w:rFonts w:ascii=\"Liberation Mono\" w:hAnsi=\"Liberation Mono\"/><w:sz w:val=\"22\"/>"
                + "</w:rPr><w:t>CCCCCCCCCC CCCCCCC</w:t><w:tab/><w:t>12.50</w:t></w:r></w:p>"));
            return $"<w:tbl><w:tblPr><w:tblW w:w=\"{gridWidth}\" w:type=\"dxa\"/>"
                + $"<w:tblLayout w:type=\"fixed\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"{gridWidth}\"/></w:tblGrid>"
                + $"<w:tr>{cell}</w:tr></w:tbl>";
        }
        static string? Tab(XElement html) => (string?)html.Descendants()
            .Single(e => e.Attribute("data-docx-tab") != null).Attribute("data-docx-tab-width");
        var bordered = Tab(Convert(Fixture(2400, true)));
        Assert.Equal(Tab(Convert(Fixture(2340, false))), bordered);
        Assert.NotEqual(Tab(Convert(Fixture(2400, false))), bordered);
    }

    [Fact]
    public void Other_border_styles_retain_the_existing_renderer()
    {
        var html = Convert(Table("<w:tr>" + Cell("CypressDouble", 8, type: "double") + Cell("CypressSingle", 2) + "</w:tr>"));
        Assert.Null(html.Descendants(Xhtml.table).Single().Attribute("data-docx-exact-table-borders"));
        Assert.Contains("double", html.ToString());
    }
}
