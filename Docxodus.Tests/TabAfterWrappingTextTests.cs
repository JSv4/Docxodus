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

    /// <summary>4680 dxa (3.25in) less Word's default 108 dxa cell margin on each side.</summary>
    private const decimal CellTextWidthInches = (4680m - 2 * 108m) / 1440m;

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
        Assert.Equal("0.500in", Style(segment)["width"]);
    }

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

    private static string Table(string firstCell)
    {
        static string Cell(string inner) => $"<w:tc><w:tcPr><w:tcW w:w='4680' w:type='dxa'/></w:tcPr>{inner}</w:tc>";
        return "<w:tbl><w:tblPr><w:tblW w:w='5000' w:type='pct'/></w:tblPr>" +
               "<w:tblGrid><w:gridCol w:w='4680'/><w:gridCol w:w='4680'/></w:tblGrid>" +
               $"<w:tr>{Cell(firstCell)}{Cell("<w:p><w:r><w:t>Second cell</w:t></w:r></w:p>")}</w:tr></w:tbl>";
    }

    private static XElement Html(string bodyXml, string extraSectionProperties = "")
    {
        using var stream = new MemoryStream();
        using (var package = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = package.AddMainDocumentPart();
            using var writer = new StreamWriter(main.GetStream(FileMode.Create));
            writer.Write($"<w:document xmlns:w='{WNs}'><w:body><w:p><w:r><w:t>Before.</w:t></w:r></w:p>{bodyXml}" +
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
