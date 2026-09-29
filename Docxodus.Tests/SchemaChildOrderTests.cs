// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// <see cref="WordprocessingMLUtil.OrderChildrenPerSchema(XElement)"/> puts the children of tables, rows,
/// paragraph properties and run properties back into schema sequence after the comparison renderer has
/// inserted or re-parented some of them (issue #837), without adding, dropping or copying anything.
/// </summary>
public class SchemaChildOrderTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace W14 = "http://schemas.microsoft.com/office/word/2010/wordml";

    private static XElement Parse(string xml) =>
        XElement.Parse($"<w:body xmlns:w='{W}' xmlns:w14='{W14}'>{xml}</w:body>");

    private static string[] ChildNames(XElement element) =>
        element.Elements().Select(e => e.Name.LocalName).ToArray();

    [Fact]
    public void RowPropertyExceptions_MoveAheadOfTheRowPropertiesAndCells()
    {
        var body = Parse("<w:tbl><w:tr><w:trPr><w:del w:id='1' w:author='A'/></w:trPr><w:tblPrEx/>" +
                         "<w:tc><w:p/></w:tc></w:tr></w:tbl>");

        Assert.True(WordprocessingMLUtil.OrderChildrenPerSchema(body));

        Assert.Equal(new[] { "tblPrEx", "trPr", "tc" }, ChildNames(body.Descendants(W + "tr").Single()));
    }

    [Fact]
    public void TableGrid_MovesAheadOfTheRows_WhileLeadingRangeMarkupStaysFirst()
    {
        var body = Parse("<w:tbl><w:bookmarkStart w:id='0' w:name='t'/><w:tblPr/><w:tr><w:tc><w:p/></w:tc></w:tr>" +
                         "<w:tblGrid/></w:tbl>");

        WordprocessingMLUtil.OrderChildrenPerSchema(body);

        Assert.Equal(new[] { "bookmarkStart", "tblPr", "tblGrid", "tr" }, ChildNames(body.Element(W + "tbl")!));
    }

    [Fact]
    public void AppendedFonts_MoveToTheirSlotInRunAndParagraphMarkProperties()
    {
        var body = Parse("<w:p><w:pPr><w:rPr><w:ins w:id='1' w:author='A'/><w:b/><w:rFonts w:ascii='Arial'/></w:rPr></w:pPr>" +
                         "<w:ins w:id='2' w:author='A'><w:r><w:rPr><w:b/><w:sz w:val='20'/><w:rFonts w:ascii='Arial'/></w:rPr>" +
                         "<w:t>x</w:t></w:r></w:ins></w:p>");

        WordprocessingMLUtil.OrderChildrenPerSchema(body);

        var rPrs = body.Descendants(W + "rPr").ToList();
        Assert.Equal(new[] { "ins", "rFonts", "b" }, ChildNames(rPrs[0]));
        Assert.Equal(new[] { "rFonts", "b", "sz" }, ChildNames(rPrs[1]));
    }

    [Fact]
    public void Word2010RunProperties_FollowTheStandardOnesAndPrecedeTheFormatChange()
    {
        var body = Parse("<w:r><w:rPr><w:b/><w:rPrChange w:id='1' w:author='A'><w:rPr/></w:rPrChange>" +
                         "<w14:ligatures w14:val='standard'/><w:lang w:val='en-US'/></w:rPr></w:r>");

        WordprocessingMLUtil.OrderChildrenPerSchema(body);

        Assert.Equal(new[] { "b", "lang", "ligatures", "rPrChange" },
            ChildNames(body.Descendants(W + "r").Single().Element(W + "rPr")!));
    }

    [Fact]
    public void UnrankedChild_TravelsWithTheRankedChildBeforeIt()
    {
        var body = Parse("<w:tbl><w:tr><w:trPr/><w:tc><w:p/></w:tc><w:bookmarkEnd w:id='0'/><w:tblPrEx/></w:tr></w:tbl>");

        WordprocessingMLUtil.OrderChildrenPerSchema(body);

        Assert.Equal(new[] { "trPr", "tc", "bookmarkEnd" },
            ChildNames(body.Descendants(W + "tr").Single()).Skip(1).ToArray());
    }

    [Fact]
    public void OrderedContent_IsLeftUntouched()
    {
        var body = Parse("<w:p><w:pPr><w:pStyle w:val='x'/><w:jc w:val='left'/><w:rPr><w:rFonts/><w:b/></w:rPr></w:pPr>" +
                         "<w:r><w:rPr><w:b/><w:sz w:val='2'/></w:rPr><w:t>x</w:t></w:r></w:p>");
        var before = body.ToString(SaveOptions.DisableFormatting);

        Assert.False(WordprocessingMLUtil.OrderChildrenPerSchema(body));
        Assert.Equal(before, body.ToString(SaveOptions.DisableFormatting));
    }

    [Fact]
    public void Reordering_KeepsEveryNode()
    {
        var body = Parse("<w:p><w:pPr><w:rPr><w:b/><w:rFonts/></w:rPr><w:jc w:val='left'/><w:pStyle w:val='x'/></w:pPr>" +
                         "<w:r><w:rPr><w:i/><w:rFonts/><!--c--><w:b/></w:rPr><w:t>x</w:t></w:r></w:p>");
        var nodes = body.DescendantNodes().Count();

        Assert.True(WordprocessingMLUtil.OrderChildrenPerSchema(body));
        Assert.Equal(nodes, body.DescendantNodes().Count());
    }
}
