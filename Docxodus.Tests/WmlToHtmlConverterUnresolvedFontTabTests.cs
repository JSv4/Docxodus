// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #847: DOCX → HTML aborted with "Sequence contains no elements" whenever a paragraph with
/// tabs contained a run whose font never resolved to a family name.
///
/// <para>The tab-width pass measures text by building a detached copy of the run and handing it to
/// the width calculator, carrying the run's (or its paragraph's) resolved font name. When neither
/// had one — no <c>w:rFonts</c> anywhere in the style chain, or theme-font references
/// (<c>w:asciiTheme</c>) the package gives no readable theme for — the calculator fell back to
/// "the enclosing paragraph's font", but the copy is detached and has no paragraph, so the lookup
/// threw and the whole conversion (and every export built on it) failed. The same fallback lived in
/// the list-marker width calculation used for right- and center-justified numbering.</para>
///
/// <para>A style chain with no <c>w:rFonts</c> at all is not a reproducer: formatting assembly
/// backfills Word's stock family for it. The failing shape is a chain that names its fonts only
/// through theme references which then do not resolve.</para>
///
/// <para>An unresolved font is not an error: the renderer emits no family and the browser applies its
/// default, so the measurement uses the same character estimate it already uses for families it
/// has no metrics for.</para>
/// </summary>
public class WmlToHtmlConverterUnresolvedFontTabTests
{
    private const string WNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace Xh = "http://www.w3.org/1999/xhtml";

    private static string Wrap(string bodyXml) =>
        $"<w:document xmlns:w=\"{WNs}\"><w:body>{bodyXml}" +
        "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/>" +
        "<w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/>" +
        "</w:sectPr></w:body></w:document>";

    private static void WritePart(OpenXmlPart part, string xml)
    {
        using var stream = part.GetStream(FileMode.Create);
        using var writer = new StreamWriter(stream, new UTF8Encoding(false));
        writer.Write(xml);
    }

    /// <summary>
    /// Builds a package from raw part XML. <paramref name="stylesXml"/> is the full styles part;
    /// <paramref name="themeXml"/> and <paramref name="numberingXml"/> are added only when given.
    /// </summary>
    private static WmlDocument Build(string bodyXml, string stylesXml, string? themeXml = null, string? numberingXml = null)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            WritePart(main, Wrap(bodyXml));
            WritePart(main.AddNewPart<StyleDefinitionsPart>(), stylesXml);
            WritePart(main.AddNewPart<DocumentSettingsPart>(),
                $"<w:settings xmlns:w=\"{WNs}\"><w:defaultTabStop w:val=\"720\"/></w:settings>");
            if (themeXml != null)
                WritePart(main.AddNewPart<ThemePart>(), themeXml);
            if (numberingXml != null)
                WritePart(main.AddNewPart<NumberingDefinitionsPart>(), numberingXml);
        }
        return new WmlDocument("unresolved-font.docx", ms.ToArray());
    }

    // A styles part that names its fonts only through theme references.
    private static readonly string StylesWithThemeFontsOnly =
        $"<w:styles xmlns:w=\"{WNs}\"><w:docDefaults><w:rPrDefault><w:rPr>" +
        "<w:rFonts w:asciiTheme=\"minorHAnsi\" w:eastAsiaTheme=\"minorHAnsi\" w:hAnsiTheme=\"minorHAnsi\" w:cstheme=\"minorBidi\"/>" +
        "<w:sz w:val=\"22\"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>" +
        "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style></w:styles>";

    // A theme part whose DrawingML elements are not in the namespace the theme-font lookup reads
    // (here the ISO 29500 strict namespace inside an otherwise transitional package), so the theme
    // references above stay unresolved even though a theme part exists.
    private const string ThemeInForeignNamespace =
        "<a:theme xmlns:a=\"http://purl.oclc.org/ooxml/drawingml/main\" name=\"T\"><a:themeElements>" +
        "<a:fontScheme name=\"T\"><a:majorFont><a:latin typeface=\"Calibri Light\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont>" +
        "<a:minorFont><a:latin typeface=\"Calibri\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont></a:fontScheme>" +
        "</a:themeElements></a:theme>";

    private const string TabbedParagraph =
        "<w:p><w:r><w:t>Name</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>Value</w:t></w:r></w:p>";

    private static XElement Convert(WmlDocument doc) =>
        WmlToHtmlConverter.ConvertToHtml(doc, new WmlToHtmlConverterSettings());

    /// <summary>The HTML spans rendered w:tab elements become (each carries its measured width).</summary>
    private static XElement[] TabSpans(XElement html) =>
        html.Descendants(Xh + "span").Where(s => s.Attribute("data-docx-tab-width") != null).ToArray();

    private static void AssertRendersTextAndSizedTab(XElement html, params string[] texts)
    {
        var body = html.Descendants(Xh + "body").Single().Value;
        foreach (var text in texts)
            Assert.Contains(text, body, StringComparison.Ordinal);
        Assert.NotEmpty(TabSpans(html));
    }

    [Fact]
    public void TabbedParagraphWithThemeFontsAndNoThemePartConverts()
    {
        var html = Convert(Build(TabbedParagraph, StylesWithThemeFontsOnly));
        AssertRendersTextAndSizedTab(html, "Name", "Value");
    }

    [Fact]
    public void TabbedParagraphWithThemeFontsAndUnreadableThemeConverts()
    {
        var html = Convert(Build(TabbedParagraph, StylesWithThemeFontsOnly, ThemeInForeignNamespace));
        AssertRendersTextAndSizedTab(html, "Name", "Value");
    }

    [Theory]
    [InlineData("end")]
    [InlineData("right")]
    [InlineData("center")]
    [InlineData("decimal")]
    public void AlignedTabStopMeasuresFollowingTextWithoutAResolvedFont(string alignment)
    {
        // Aligned stops measure the text AFTER the tab to place it, through a separate detached
        // copy of the run — the path a table-of-contents page-number column takes.
        var body =
            $"<w:p><w:pPr><w:tabs><w:tab w:val=\"{alignment}\" w:leader=\"dot\" w:pos=\"8640\"/></w:tabs></w:pPr>" +
            "<w:r><w:t>Chapter</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>12.5</w:t></w:r></w:p>";

        var html = Convert(Build(body, StylesWithThemeFontsOnly));

        AssertRendersTextAndSizedTab(html, "Chapter", "12.5");
        // Text placed against a stop 6 inches in cannot leave a zero-width tab ahead of it.
        var widths = TabSpans(html)
            .Select(s => decimal.Parse((string)s.Attribute("data-docx-tab-width")!, System.Globalization.CultureInfo.InvariantCulture))
            .ToArray();
        Assert.Contains(widths, w => w > 1m);
    }

    [Theory]
    [InlineData("right")]
    [InlineData("center")]
    public void JustifiedListMarkerWithoutAResolvedFontConverts(string justification)
    {
        // Right- and center-justified list levels measure the generated marker to widen the
        // hanging indent; the marker is a detached run carrying the paragraph's resolved font.
        var numbering =
            $"<w:numbering xmlns:w=\"{WNs}\"><w:abstractNum w:abstractNumId=\"0\">" +
            $"<w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.\"/><w:lvlJc w:val=\"{justification}\"/>" +
            "<w:pPr><w:ind w:left=\"720\"/></w:pPr></w:lvl></w:abstractNum>" +
            "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num></w:numbering>";
        var body =
            "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr>" +
            "<w:r><w:t>First item</w:t></w:r></w:p>";

        var html = Convert(Build(body, StylesWithThemeFontsOnly, numberingXml: numbering));

        var text = html.Descendants(Xh + "body").Single().Value;
        Assert.Contains("1.", text, StringComparison.Ordinal);
        Assert.Contains("First item", text, StringComparison.Ordinal);
    }
}
