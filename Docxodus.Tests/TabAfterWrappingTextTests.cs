// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Text before a tab that is longer than its line wraps, and the tab is measured from the pen on the
/// last line, as Word lays it out (issue #891). The converter used to measure that text as one
/// unwrapped line and pin it in a no-wrap box as wide as the tab stop it chose — 9 inches in a
/// 3.25-inch cell — which forced the table off the page.
/// </summary>
public class TabAfterWrappingTextTests
{
    private const string WNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private const string Sentence =
        "the parties agree that the supplier shall deliver the services described in the schedule within thirty days";

    /// <summary>4680 dxa (3.25in): the table states no cell margins and the document has no table
    /// styles, so the cell has none (issue #1219).</summary>
    private const decimal CellTextWidthInches = 4680m / 1440m;

    /// <summary>US Letter with one-inch margins.</summary>
    private const decimal BodyTextWidthInches = 6.5m;

    [Fact]
    public void TableCell_LongTextBeforeTab_IsNotPinnedWiderThanTheCell()
    {
        var html = Html(Table(TabParagraph(Sentence)));

        AssertNoBoxWiderThan(html, CellTextWidthInches);
        Assert.Contains(Sentence, html.Value);
        Assert.Contains("12.50", html.Value);
    }

    [Fact]
    public void TableCell_LongTextBeforeTab_WrapsInsideTheCellsStatedMargins()
    {
        // One-inch margins each side leave 1800 dxa of the 4680-dxa cell for text (issue #1219): the
        // tab lands where it does in a 1800-dxa cell with no margins, not where it does in the wide cell.
        static decimal TabWidth(XElement html) => decimal.Parse(
            (string)html.Descendants().Single(e => e.Attribute("data-docx-tab") != null).Attribute("data-docx-tab-width")!,
            CultureInfo.InvariantCulture);
        var inset = TabWidth(Html(Table(TabParagraph(Sentence),
            "<w:tblCellMar><w:left w:w='1440' w:type='dxa'/><w:right w:w='1440' w:type='dxa'/></w:tblCellMar>")));

        Assert.Equal(TabWidth(Html(Table(TabParagraph(Sentence), cellWidth: 1800))), inset);
        Assert.NotEqual(TabWidth(Html(Table(TabParagraph(Sentence)))), inset);
    }

    [Fact]
    public void Body_LongTextBeforeTab_IsNotPinnedWiderThanTheTextColumn()
    {
        var html = Html(TabParagraph(Sentence + " and " + Sentence));

        AssertNoBoxWiderThan(html, BodyTextWidthInches);
    }

    [Fact]
    public void TwoColumnSection_TextWiderThanAColumn_WrapsAtTheColumnNotThePage()
    {
        // About 4.5in of text: it fits the 6.5in page text width but not one of two 3.0in columns.
        var html = Html(TabParagraph("the parties agree that the supplier shall deliver the services"),
            "<w:cols w:num='2' w:space='720'/>");

        Assert.DoesNotContain(html.Descendants(), e => Style(e).GetValueOrDefault("display") == "inline-flex");
        AssertNoBoxWiderThan(html, 3.0m);
    }

    [Fact]
    public void TableCell_LongTextBeforeTab_TabAdvancesToTheNextStopOnTheLastLine()
    {
        var html = Html(Table(TabParagraph(Sentence)));

        // Default stops every half inch: measured from the pen on the wrapped last line, the
        // advance is at most one stop interval and never reaches past the cell.
        var tab = Assert.Single(html.Descendants(), e => e.Attribute("data-docx-tab") != null);
        var width = decimal.Parse((string)tab.Attribute("data-docx-tab-width")!, CultureInfo.InvariantCulture);
        Assert.InRange(width, 0m, 0.5m);
    }

    [Fact]
    public void TableCell_ShortTextBeforeTab_KeepsThePinnedTabSegment()
    {
        var html = Html(Table(TabParagraph("Total")));

        var segment = Assert.Single(html.Descendants(Xhtml("span")),
            e => Style(e).GetValueOrDefault("display") == "inline-flex");
        Assert.Equal("36pt", Style(segment)["width"]);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TextBox_LongPrefixWrapsInsideItsOwnWidth(bool drawing)
    {
        const string prefix = "Cedar trainees catalog amber pebbles while quiet swallows circle the orchard and return beside the fountain";
        var html = Html(TextBox(TabParagraph(prefix), drawing));
        var box = Assert.Single(html.Descendants(), e => Style(e).GetValueOrDefault("display") == "inline-block" && Style(e).GetValueOrDefault("height") == "360pt");
        Assert.Contains(prefix, box.Value);
        Assert.Contains("12.50", box.Value);
        Assert.DoesNotContain(box.Descendants(), e => Style(e).GetValueOrDefault("display") == "inline-flex");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TextBox_FittingPrefixKeepsItsAlignedTab(bool drawing)
    {
        var html = Html(TextBox(TabParagraph("Cedar"), drawing));
        var box = Assert.Single(html.Descendants(), e => Style(e).GetValueOrDefault("display") == "inline-block" && Style(e).GetValueOrDefault("height") == "360pt");
        Assert.Contains(box.Descendants(), e => Style(e).GetValueOrDefault("display") == "inline-flex");
    }

    [Fact]
    public void DrawingTextBox_WrapMeasurementExcludesTheSourceInsets()
    {
        const string prefix = "Cedar trainees catalog amber pebbles while quiet swallows circle the orchard";
        static string? TabWidth(XElement html) => (string?)html.Descendants()
            .Single(e => e.Attribute("data-docx-tab") != null).Attribute("data-docx-tab-width");
        var inset = Html(TextBox(TabParagraph(prefix), true, 192, 24));
        var narrow = Html(TextBox(TabParagraph(prefix), true, 144, 0));
        var wide = Html(TextBox(TabParagraph(prefix), true, 192, 0));
        Assert.Equal(TabWidth(narrow), TabWidth(inset));
        Assert.NotEqual(TabWidth(wide), TabWidth(inset));
    }

    [Theory]
    [InlineData(0, false)]
    [InlineData(50000, true)]
    public void DrawingTextBox_ActiveRelativeWidthRetainsUnknownWrapWidth(int percentage, bool unknown)
    {
        const string prefix = "Cedar trainees catalog amber pebbles while quiet swallows circle the orchard";
        var source = TextBox(TabParagraph(prefix), true)
            .Replace("<wp:inline>", "<wp:anchor>")
            .Replace("</wp:inline>", $"<wp14:sizeRelH relativeFrom=\"margin\"><wp14:pctWidth>{percentage}</wp14:pctWidth></wp14:sizeRelH></wp:anchor>");
        var html = Html(source);
        Assert.Equal(unknown, html.Descendants().Any(e => Style(e).GetValueOrDefault("display") == "inline-flex"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DrawingTextBox_NegativeCssInsetsDoNotEnlargeTheWrapWidth(bool border)
    {
        const string prefix = "Cedar trainees catalog amber pebbles while quiet swallows circle the orchard";
        var ordinary = TextBox(TabParagraph(prefix), true);
        var negative = border
            ? ordinary.Replace("<wps:txbx>", "<wps:spPr><a:ln w=\"-304800\"><a:solidFill><a:srgbClr val=\"225566\"/></a:solidFill></a:ln></wps:spPr><wps:txbx>")
            : TextBox(TabParagraph(prefix), true, 144, -24);
        static string? TabWidth(XElement html) => (string?)html.Descendants()
            .Single(e => e.Attribute("data-docx-tab") != null).Attribute("data-docx-tab-width");
        Assert.Equal(TabWidth(Html(ordinary)), TabWidth(Html(negative)));
    }

    private static string TextBox(string paragraph, bool drawing, int width = 144, int inset = 0) => drawing
        ? $"""
          <w:p><w:r><w:drawing><wp:inline><wp:extent cx="{width * 12700}" cy="4572000"/>
          <wp:docPr id="19" name="Cedar tab frame"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp><wps:txbx><w:txbxContent>{paragraph}</w:txbxContent></wps:txbx>
          <wps:bodyPr lIns="{inset * 12700}" rIns="{inset * 12700}" tIns="0" bIns="0"/></wps:wsp>
          </a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
          """
        : $"""
          <w:p><w:r><w:pict><v:shape id="CedarFrame" style="width:{width}pt;height:360pt">
          <v:textbox inset="0,0,0,0"><w:txbxContent>{paragraph}</w:txbxContent></v:textbox>
          </v:shape></w:pict></w:r></w:p>
          """;

    private static void AssertNoBoxWiderThan(XElement html, decimal inches)
    {
        foreach (var element in html.Descendants())
        {
            var style = Style(element);
            foreach (var property in new[] { "width", "min-width" })
            {
                if (style.TryGetValue(property, out var value) && value.EndsWith("in"))
                {
                    var measured = decimal.Parse(value[..^2], CultureInfo.InvariantCulture);
                    Assert.True(measured <= inches,
                        $"<{element.Name.LocalName}> declares {property}: {value}, wider than the {inches:0.000}in line");
                }
            }
        }
    }

    private static Dictionary<string, string> Style(XElement element) =>
        ((string?)element.Attribute("style") ?? string.Empty)
            .Split(';', System.StringSplitOptions.RemoveEmptyEntries)
            .Select(d => d.Split(':', 2))
            .Where(p => p.Length == 2)
            .GroupBy(p => p[0].Trim())
            .ToDictionary(g => g.Key, g => g.Last()[1].Trim());

    private static XName Xhtml(string name) => XName.Get(name, "http://www.w3.org/1999/xhtml");

    private static string TabParagraph(string before) =>
        $"<w:p><w:r><w:t>{before}</w:t></w:r><w:r><w:tab/><w:t>12.50</w:t></w:r></w:p>";

    private static string Table(string firstCell, string tableProperties = "", int cellWidth = 4680)
    {
        string Cell(string inner) => $"<w:tc><w:tcPr><w:tcW w:w='{cellWidth}' w:type='dxa'/></w:tcPr>{inner}</w:tc>";
        return $"<w:tbl><w:tblPr><w:tblW w:w='5000' w:type='pct'/>{tableProperties}</w:tblPr>" +
               $"<w:tblGrid><w:gridCol w:w='{cellWidth}'/><w:gridCol w:w='{cellWidth}'/></w:tblGrid>" +
               $"<w:tr>{Cell(firstCell)}{Cell("<w:p><w:r><w:t>Second cell</w:t></w:r></w:p>")}</w:tr></w:tbl>";
    }

    private static XElement Html(string bodyXml, string extraSectionProperties = "")
    {
        using var stream = new MemoryStream();
        using (var package = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = package.AddMainDocumentPart();
            using var writer = new StreamWriter(main.GetStream(FileMode.Create));
            writer.Write($"<w:document xmlns:w='{WNs}' xmlns:wp='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing' " +
                         "xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main' xmlns:wps='http://schemas.microsoft.com/office/word/2010/wordprocessingShape' " +
                         $"xmlns:v='urn:schemas-microsoft-com:vml' xmlns:wp14='http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing'><w:body><w:p><w:r><w:t>Before.</w:t></w:r></w:p>{bodyXml}" +
                         "<w:sectPr><w:pgSz w:w='12240' w:h='15840'/>" +
                         "<w:pgMar w:top='1440' w:right='1440' w:bottom='1440' w:left='1440' w:header='720' w:footer='720' w:gutter='0'/>" +
                         extraSectionProperties + "</w:sectPr></w:body></w:document>");
        }
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("tab.docx", stream.ToArray()),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false });
        // Make the namespace of the generated tree explicit for the queries above.
        return XElement.Parse(html.ToString());
    }
}
