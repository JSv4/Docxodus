using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Tab layout measures the text around each tab with a DETACHED copy of the run. When neither the run
/// nor its paragraph resolved a concrete font name — here, fonts named only through theme references in
/// a package with no theme part — the measurement looked for the run's paragraph on that detached copy
/// and threw <c>Sequence contains no elements</c>, failing the whole conversion (issue #847). One
/// unmeasurable run must never fail the document; it measures with the character-width estimate the
/// converter already uses for an unknown font.
/// </summary>
public class WmlToHtmlRunWidthFallbackTests
{
    private const string WNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>Default run fonts named only by theme slot, and no theme part to resolve them.</summary>
    private const string ThemeOnlyStyles =
        "<w:styles xmlns:w=\"" + WNs + "\"><w:docDefaults><w:rPrDefault><w:rPr>" +
        "<w:rFonts w:asciiTheme=\"minorHAnsi\" w:hAnsiTheme=\"minorHAnsi\"/>" +
        "</w:rPr></w:rPrDefault></w:docDefaults></w:styles>";

    /// <summary>One decimal list level whose number is justified <paramref name="lvlJc"/>; right and
    /// center justification make the formatting pass measure the marker to widen the hanging indent.</summary>
    private static string DecimalNumbering(string lvlJc) =>
        "<w:numbering xmlns:w=\"" + WNs + "\"><w:abstractNum w:abstractNumId=\"0\"><w:lvl w:ilvl=\"0\">" +
        $"<w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.\"/><w:lvlJc w:val=\"{lvlJc}\"/>" +
        "<w:pPr><w:ind w:left=\"720\" w:hanging=\"360\"/></w:pPr></w:lvl></w:abstractNum>" +
        "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num></w:numbering>";

    private static byte[] BuildDoc(string bodyInner, string? numberingJc = null)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            Write(main, $"<w:document xmlns:w=\"{WNs}\"><w:body>{bodyInner}</w:body></w:document>");
            Write(main.AddNewPart<StyleDefinitionsPart>(), ThemeOnlyStyles);
            if (numberingJc is not null)
                Write(main.AddNewPart<NumberingDefinitionsPart>(), DecimalNumbering(numberingJc));
        }
        return ms.ToArray();
    }

    private static void Write(OpenXmlPart part, string xml)
    {
        using var s = part.GetStream(FileMode.Create);
        using var w = new StreamWriter(s);
        w.Write(xml);
    }

    private static string TabbedParagraph(string? tabVal)
    {
        var tabs = tabVal is null
            ? string.Empty
            : $"<w:pPr><w:tabs><w:tab w:val=\"{tabVal}\" w:pos=\"6000\"/></w:tabs></w:pPr>";
        return $"<w:p>{tabs}<w:r><w:t>Total</w:t></w:r><w:r><w:tab/><w:t>12.50</w:t></w:r></w:p>";
    }

    private static string Convert(byte[] docx) =>
        WmlToHtmlConverter.ConvertToHtml(new WmlDocument("run-width.docx", docx), new WmlToHtmlConverterSettings())
            .ToString(SaveOptions.DisableFormatting);

    [Theory]
    [InlineData(null)] // default tab stops
    [InlineData("left")]
    [InlineData("right")]
    [InlineData("center")]
    [InlineData("decimal")]
    public void TabbedParagraph_WithNoResolvableFont_Converts(string? tabVal)
    {
        var html = Convert(BuildDoc(TabbedParagraph(tabVal)));

        Assert.Contains("Total", html);
        Assert.Contains("12.50", html);
    }

    [Theory]
    [InlineData("left")]
    [InlineData("right")]
    [InlineData("center")]
    public void ListItem_WithNoResolvableFont_ConvertsWithItsMarker(string lvlJc)
    {
        var docx = BuildDoc(
            "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr>" +
            "<w:r><w:t>First item</w:t></w:r></w:p>",
            numberingJc: lvlJc);

        var html = XElement.Parse(Convert(docx));

        // The generated list-number run renders its label as its own text node.
        Assert.Contains(html.DescendantNodes().OfType<XText>(), t => t.Value.Trim() == "1.");
        Assert.Contains("First item", html.Value);
    }

    [Theory]
    [InlineData(0)] // flowing HTML
    [InlineData(1)] // paginated view, the path the PDF export renders
    public void SharedConversionFacade_WithNoResolvableFont_DoesNotFail(int paginationMode)
    {
        var docx = BuildDoc(TabbedParagraph("right"));

        var html = HtmlConversionOps.ConvertToHtml(docx, new HtmlConversionOptions { PaginationMode = paginationMode });

        Assert.Contains("12.50", html);
    }
}
