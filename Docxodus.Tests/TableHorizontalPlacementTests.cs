using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Where Word puts a table's first cell text across the line (issue #1219). Word for the web
/// measurements are in docs/ooxml_corner_cases.md: a cell with no margin stated anywhere has none,
/// a table that names no style takes the default table style's, and before compatibility mode 15
/// Word aligns the first cell's text, not the table edge, with a stated indent.
/// </summary>
public class TableHorizontalPlacementTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string Margins108 =
        "<w:tblCellMar><w:left w:w=\"108\" w:type=\"dxa\"/><w:right w:w=\"108\" w:type=\"dxa\"/></w:tblCellMar>";
    private const string DefaultTableStyle =
        "<w:style w:type=\"table\" w:default=\"1\" w:styleId=\"TableNormal\"><w:name w:val=\"Normal Table\"/><w:tblPr>"
        + "<w:tblInd w:w=\"0\" w:type=\"dxa\"/>" + Margins108 + "</w:tblPr></w:style>";

    [Fact]
    public void Cell_has_no_margin_when_nothing_states_one()
    {
        var html = Convert(Table(), mode: "15");
        Assert.Equal("0pt", Cell(html)["padding-left"]);
        Assert.Equal("0pt", Cell(html)["padding-right"]);
    }

    [Fact]
    public void Stated_table_cell_margins_still_apply()
    {
        var html = Convert(Table(props: Margins108), mode: "15");
        Assert.Equal("5.4pt", Cell(html)["padding-left"]);
    }

    [Fact]
    public void Table_that_names_no_style_takes_the_default_table_styles_cell_margins()
    {
        var html = Convert(Table(), mode: "15", styles: DefaultTableStyle);
        Assert.Equal("5.4pt", Cell(html)["padding-left"]);
        Assert.Equal("5.4pt", Cell(html)["padding-right"]);
        Assert.Equal("0", Table(html)["margin-left"]); // mode 15: the table edge sits at the indent
    }

    [Fact]
    public void Table_margins_beat_the_default_table_style()
    {
        var html = Convert(Table(props: "<w:tblCellMar><w:left w:w=\"288\" w:type=\"dxa\"/></w:tblCellMar>"),
            mode: "15", styles: DefaultTableStyle);
        Assert.Equal("14.4pt", Cell(html)["padding-left"]);
        Assert.Equal("5.4pt", Cell(html)["padding-right"]);
    }

    [Fact]
    public void Default_table_style_margin_does_not_override_the_tables_other_spelling_of_the_side()
    {
        var styles = DefaultTableStyle.Replace("<w:left ", "<w:start ");
        var html = Convert(Table(props: "<w:tblCellMar><w:left w:w=\"288\" w:type=\"dxa\"/></w:tblCellMar>"),
            mode: "15", styles: styles);
        Assert.Equal("14.4pt", Cell(html)["padding-left"]);
    }

    [Theory]
    [InlineData("14")]
    [InlineData("12")]
    [InlineData(null)]
    public void Before_mode_15_a_stated_indent_places_the_first_cells_text(string? mode)
    {
        // Out by the 5.4pt margin and half the 1pt border: the text lands on the indent.
        Assert.Equal("-5.9pt", Table(Convert(Table(props: "<w:tblInd w:w=\"0\" w:type=\"dxa\"/>" + Margins108), mode))["margin-left"]);
        Assert.Equal("30.1pt", Table(Convert(Table(props: "<w:tblInd w:w=\"720\" w:type=\"dxa\"/>" + Margins108), mode))["margin-left"]);
    }

    [Fact]
    public void Before_mode_15_the_first_cells_own_margin_sets_the_pull()
    {
        var html = Convert(Table(props: "<w:tblInd w:w=\"0\" w:type=\"dxa\"/>" + Margins108,
            cellProps: "<w:tcMar><w:left w:w=\"288\" w:type=\"dxa\"/></w:tcMar>"), mode: "14");
        Assert.Equal("-14.9pt", Table(html)["margin-left"]);
    }

    [Fact]
    public void Before_mode_15_the_default_table_style_pulls_the_table_out_too()
    {
        var html = Convert(Table(), mode: "14", styles: DefaultTableStyle);
        Assert.Equal("-5.9pt", Table(html)["margin-left"]);
    }

    [Fact]
    public void Before_mode_15_an_unstyled_table_with_no_indent_centres_its_border_on_the_margin()
    {
        var html = Convert(Table(props: Margins108), mode: "14");
        Assert.Equal("-0.5pt", Table(html)["margin-left"]);
        Assert.Equal("5.4pt", Cell(html)["padding-left"]);
    }

    [Fact]
    public void Mode_15_keeps_the_table_edge_at_the_stated_indent()
    {
        var html = Convert(Table(props: "<w:tblInd w:w=\"720\" w:type=\"dxa\"/>" + Margins108), mode: "15");
        Assert.Equal("36pt", Table(html)["margin-left"]);
    }

    [Fact]
    public void Floating_table_is_not_pulled_out()
    {
        var html = Convert(Table(props: "<w:tblpPr w:leftFromText=\"180\" w:vertAnchor=\"text\" w:tblpX=\"1440\"/>"
            + "<w:tblInd w:w=\"0\" w:type=\"dxa\"/>" + Margins108), mode: "14");
        Assert.Equal("0", Table(html)["margin-left"]);
    }

    // A one-cell table with 1pt borders; props and cellProps go into its w:tblPr and w:tcPr.
    private static string Table(string props = "", string cellProps = "") =>
        "<w:tbl><w:tblPr><w:tblW w:w=\"5000\" w:type=\"dxa\"/>" + props + "<w:tblBorders>"
        + string.Concat(new[] { "top", "left", "bottom", "right" }.Select(s => $"<w:{s} w:val=\"single\" w:sz=\"8\"/>"))
        + "</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w=\"5000\"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w=\"5000\" w:type=\"dxa\"/>"
        + cellProps + "</w:tcPr><w:p><w:r><w:t>Juniper cell.</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";

    private static XElement Convert(string body, string? mode, string styles = "")
    {
        var compat = mode == null ? "" : "<w:compat><w:compatSetting w:name=\"compatibilityMode\" "
            + $"w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"{mode}\"/></w:compat>";
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            Feed(main, $"<w:document xmlns:w=\"{W}\"><w:body><w:p><w:r><w:t>Before.</w:t></w:r></w:p>{body}"
                + "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" "
                + "w:left=\"1440\" w:header=\"720\" w:footer=\"720\"/></w:sectPr></w:body></w:document>");
            Feed(main.AddNewPart<StyleDefinitionsPart>(), $"<w:styles xmlns:w=\"{W}\">{styles}</w:styles>");
            Feed(main.AddNewPart<DocumentSettingsPart>(), $"<w:settings xmlns:w=\"{W}\">{compat}</w:settings>");
        }
        return WmlToHtmlConverter.ConvertToHtml(new WmlDocument("Juniper.docx", stream.ToArray()),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false });
    }

    private static void Feed(OpenXmlPart part, string xml)
    {
        using var data = new MemoryStream(Encoding.UTF8.GetBytes(xml));
        part.FeedData(data);
    }

    private static Dictionary<string, string> Table(XElement html) =>
        Style(html.Descendants().Single(e => e.Name.LocalName == "table"));

    private static Dictionary<string, string> Cell(XElement html) =>
        Style(html.Descendants().Single(e => e.Name.LocalName == "td"));

    private static Dictionary<string, string> Style(XElement element) =>
        ((string?)element.Attribute("style") ?? "").Split(';', System.StringSplitOptions.RemoveEmptyEntries)
            .Select(d => d.Split(':', 2)).ToDictionary(d => d[0].Trim(), d => d[1].Trim());
}
