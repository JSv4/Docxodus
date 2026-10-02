// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// ISO/IEC 29500 names a paragraph's leading and trailing indents <c>w:start</c> and <c>w:end</c>;
/// the transitional schema also spells them <c>w:left</c> and <c>w:right</c>. Word writes the
/// transitional spelling and LibreOffice writes the other, and the converter read only
/// <c>w:left</c>/<c>w:right</c>: a LibreOffice-saved document lost every indent, and its lists drew
/// the marker past the left edge of the text, because the hanging indent still applied (issue #894).
/// When one <c>w:ind</c> carries both spellings, <c>w:start</c>/<c>w:end</c> win.
/// </summary>
public class WmlIndStartEndTests
{
    private const string WNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>A paragraph style <c>Indented</c> whose <c>w:ind</c> is <paramref name="styleInd"/>.</summary>
    private static string Styles(string styleInd) =>
        $"<w:styles xmlns:w=\"{WNs}\">" +
        "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>" +
        "<w:style w:type=\"paragraph\" w:styleId=\"Indented\"><w:name w:val=\"Indented\"/>" +
        $"<w:basedOn w:val=\"Normal\"/><w:pPr>{styleInd}</w:pPr></w:style></w:styles>";

    /// <summary>One decimal list level whose <c>w:ind</c> is <paramref name="lvlInd"/>.</summary>
    private static string Numbering(string lvlInd) =>
        $"<w:numbering xmlns:w=\"{WNs}\"><w:abstractNum w:abstractNumId=\"0\"><w:lvl w:ilvl=\"0\">" +
        "<w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.\"/><w:lvlJc w:val=\"left\"/>" +
        $"<w:pPr>{lvlInd}</w:pPr></w:lvl></w:abstractNum>" +
        "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num></w:numbering>";

    private static byte[] BuildDoc(string bodyInner, string styleInd = "", string lvlInd = "")
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            Write(main, $"<w:document xmlns:w=\"{WNs}\"><w:body>{bodyInner}</w:body></w:document>");
            Write(main.AddNewPart<StyleDefinitionsPart>(), Styles(styleInd));
            Write(main.AddNewPart<DocumentSettingsPart>(), $"<w:settings xmlns:w=\"{WNs}\"/>");
            Write(main.AddNewPart<NumberingDefinitionsPart>(), Numbering(lvlInd));
        }
        return ms.ToArray();
    }

    private static void Write(OpenXmlPart part, string xml)
    {
        using var s = part.GetStream(FileMode.Create);
        using var w = new StreamWriter(s);
        w.Write(xml);
    }

    private static string Paragraph(string text, string pPr = "") =>
        $"<w:p><w:pPr>{pPr}</w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>";

    private const string ListItem = "<w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr>";

    private static XElement Convert(byte[] docx) =>
        WmlToHtmlConverter.ConvertToHtml(
            new WmlDocument("ind.docx", docx),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false });

    /// <summary>The inline style of the innermost element whose own text includes <paramref name="text"/>.</summary>
    private static Dictionary<string, string> StyleOf(XElement html, string text, string elementName = "p")
    {
        var element = html.Descendants()
            .Where(e => e.Name.LocalName == elementName && e.Value.Contains(text))
            .Last();
        return ((string?)element.Attribute("style") ?? string.Empty)
            .Split(';', System.StringSplitOptions.RemoveEmptyEntries)
            .Select(d => d.Split(':', 2))
            .ToDictionary(kv => kv[0].Trim(), kv => kv[1].Trim());
    }

    [Theory]
    [InlineData("left")]
    [InlineData("start")]
    public void LeadingIndent_SetsTheLeftMargin(string leading)
    {
        var html = Convert(BuildDoc(Paragraph("Indented one inch.", $"<w:ind w:{leading}=\"1440\"/>")));

        Assert.Equal("1.00in", StyleOf(html, "Indented one inch.")["margin-left"]);
    }

    [Theory]
    [InlineData("right")]
    [InlineData("end")]
    public void TrailingIndent_SetsTheRightMargin(string trailing)
    {
        var html = Convert(BuildDoc(Paragraph("Indented on the right.", $"<w:ind w:{trailing}=\"1440\"/>")));

        Assert.Equal("1.00in", StyleOf(html, "Indented on the right.")["margin-right"]);
    }

    [Theory]
    [InlineData("left")]
    [InlineData("start")]
    public void ListLevelIndent_KeepsTheMarkerInsideTheContentBox(string leading)
    {
        var html = Convert(BuildDoc(
            Paragraph("First list item.", ListItem),
            lvlInd: $"<w:ind w:{leading}=\"720\" w:hanging=\"360\"/>"));

        var style = StyleOf(html, "First list item.");
        Assert.Equal("0.50in", style["margin-left"]);
        Assert.Equal("-0.25in", style["text-indent"]);
    }

    [Theory]
    [InlineData("left", "start")]
    [InlineData("start", "left")]
    public void ParagraphIndent_OverridesItsStyle_InEitherSpelling(string styleSpelling, string paragraphSpelling)
    {
        var html = Convert(BuildDoc(
            Paragraph("Overridden.", $"<w:pStyle w:val=\"Indented\"/><w:ind w:{paragraphSpelling}=\"720\"/>"),
            styleInd: $"<w:ind w:{styleSpelling}=\"2880\"/>"));

        Assert.Equal("0.50in", StyleOf(html, "Overridden.")["margin-left"]);
    }

    [Fact]
    public void StartAndEnd_Win_WhenOneIndCarriesBothSpellings()
    {
        var html = Convert(BuildDoc(Paragraph(
            "Both spellings.",
            "<w:ind w:left=\"2880\" w:start=\"720\" w:right=\"2880\" w:end=\"720\"/>")));

        var style = StyleOf(html, "Both spellings.");
        Assert.Equal("0.50in", style["margin-left"]);
        Assert.Equal("0.50in", style["margin-right"]);
    }

    [Theory]
    [InlineData("left")]
    [InlineData("start")]
    public void BorderedParagraphGroup_TakesTheLeadingIndent(string leading)
    {
        var html = Convert(BuildDoc(Paragraph(
            "Boxed.",
            "<w:pBdr><w:top w:val=\"single\" w:sz=\"4\" w:space=\"1\" w:color=\"auto\"/></w:pBdr>" +
            $"<w:ind w:{leading}=\"1440\"/>")));

        Assert.Equal("1.00in", StyleOf(html, "Boxed.", "div")["margin-left"]);
    }

    /// <summary>
    /// One document touching every place the converter reads the indent: paragraph margins, a list
    /// level, a bordered group, and tab layout (a hanging indent puts the first tab stop at the
    /// leading indent, so the tab's rendered width depends on it). The two spellings must convert
    /// to identical HTML.
    /// </summary>
    [Fact]
    public void StrictSpelling_ConvertsExactlyLikeTheTransitionalSpelling()
    {
        static byte[] Doc(string leading, string trailing) => BuildDoc(
            Paragraph("Indented.", $"<w:ind w:{leading}=\"1440\" w:{trailing}=\"720\"/>") +
            Paragraph("Item.", ListItem) +
            Paragraph("Boxed.",
                "<w:pBdr><w:top w:val=\"single\" w:sz=\"4\" w:space=\"1\" w:color=\"auto\"/></w:pBdr>" +
                $"<w:ind w:{leading}=\"1440\"/>") +
            $"<w:p><w:pPr><w:ind w:{leading}=\"1440\" w:hanging=\"1440\"/></w:pPr>" +
            "<w:r><w:t>Term</w:t></w:r><w:r><w:tab/><w:t>Definition</w:t></w:r></w:p>",
            styleInd: $"<w:ind w:{leading}=\"2880\"/>",
            lvlInd: $"<w:ind w:{leading}=\"720\" w:hanging=\"360\"/>");

        var transitional = Convert(Doc("left", "right")).ToString(SaveOptions.DisableFormatting);
        var strict = Convert(Doc("start", "end")).ToString(SaveOptions.DisableFormatting);

        Assert.Equal(transitional, strict);
    }

    [Theory]
    [InlineData("left")]
    [InlineData("start")]
    public void FormattingAssembler_PutsTheListTabStopAtTheLeadingIndent(string leading)
    {
        var assembled = FormattingAssembler.AssembleFormatting(
            new WmlDocument("ind.docx", BuildDoc(
                Paragraph("First list item.", ListItem),
                lvlInd: $"<w:ind w:{leading}=\"720\" w:hanging=\"360\"/>")),
            new FormattingAssemblerSettings());

        using var ms = new MemoryStream();
        ms.Write(assembled.DocumentByteArray, 0, assembled.DocumentByteArray.Length);
        using var doc = WordprocessingDocument.Open(ms, false);
        var pPr = doc.MainDocumentPart!.GetXDocument().Descendants(W.p).Single().Element(W.pPr)!;

        var tabStops = pPr.Elements(W.tabs).Elements(W.tab).Select(t => (string?)t.Attribute(W.pos));
        Assert.Contains("720", tabStops);
    }

    [Theory]
    [InlineData("left")]
    [InlineData("start")]
    public void Session_IndentDelta_AdjustsTheSpellingTheParagraphUses(string leading)
    {
        using var s = new DocxSession(BuildDoc(Paragraph("Indented.", $"<w:ind w:{leading}=\"720\"/>")));
        var anchor = s.Project().AnchorIndex.Keys.First(k => k.StartsWith("p:"));

        var r = s.SetParagraphFormat(anchor, new ParagraphFormatOp { IndentDelta = 360 });

        Assert.True(r.Success, r.Error?.Message);
        var ind = XElement.Parse(s.Raw.GetXml(anchor)).Descendants(W.ind).Single();
        Assert.Equal("1080", (string?)ind.Attribute(W.w + leading));
        Assert.Single(ind.Attributes(), a => a.Name == W.left || a.Name == W.start);
        Assert.Equal(1080, s.GetFormatting(anchor)!.DirectParagraph.LeftIndentTwips);
    }

    [Theory]
    [InlineData("left", "start")]
    [InlineData("start", "left")]
    public void Session_EffectiveIndent_TakesTheParagraphOverride_InEitherSpelling(string styleSpelling, string paragraphSpelling)
    {
        using var s = new DocxSession(BuildDoc(
            Paragraph("Overridden.", $"<w:pStyle w:val=\"Indented\"/><w:ind w:{paragraphSpelling}=\"720\"/>"),
            styleInd: $"<w:ind w:{styleSpelling}=\"2880\"/>"));
        var anchor = s.Project().AnchorIndex.Keys.First(k => k.StartsWith("p:"));

        Assert.Equal(720, s.GetFormatting(anchor)!.EffectiveParagraph.LeftIndentTwips);
    }
}
