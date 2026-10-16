using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #1151: Word keeps a paragraph's spacing after when the next paragraph belongs to the same
/// numbered list. Only <c>w:contextualSpacing</c> removes the spacing between paragraphs, and only
/// between paragraphs of the same style. The converter used to drop the spacing after of every
/// numbered paragraph but the last in a run of the same list, so a numbered legal document lost
/// 12pt between each of its sections.
/// </summary>
public class NumberedParagraphSpacingTests
{
    private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static string Item(string text, string style = "Clause") =>
        $"<w:p><w:pPr><w:pStyle w:val=\"{style}\"/><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr>"
        + $"<w:r><w:t>{text}</w:t></w:r></w:p>";

    private static byte[] Docx(string body, bool contextualSpacing, string contextual = "<w:contextualSpacing/>")
    {
        contextual = contextualSpacing ? contextual : string.Empty;
        var styles =
            $"<w:styles xmlns:w=\"{WordNamespace}\">"
            + "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>"
            + "<w:style w:type=\"paragraph\" w:styleId=\"Clause\"><w:name w:val=\"Clause\"/><w:basedOn w:val=\"Normal\"/>"
            + $"<w:pPr><w:spacing w:after=\"240\"/>{contextual}</w:pPr></w:style>"
            + "</w:styles>";
        var numbering =
            $"<w:numbering xmlns:w=\"{WordNamespace}\">"
            + "<w:abstractNum w:abstractNumId=\"0\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/>"
            + "<w:lvlText w:val=\"%1.\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"720\" w:hanging=\"360\"/></w:pPr></w:lvl>"
            + "</w:abstractNum><w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num></w:numbering>";
        var document = $"<w:document xmlns:w=\"{WordNamespace}\"><w:body>{body}</w:body></w:document>";

        using var stream = new MemoryStream();
        using (var package = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = package.AddMainDocumentPart();
            Write(main, document);
            Write(main.AddNewPart<StyleDefinitionsPart>(), styles);
            Write(main.AddNewPart<NumberingDefinitionsPart>(), numbering);
            Write(main.AddNewPart<DocumentSettingsPart>(), $"<w:settings xmlns:w=\"{WordNamespace}\"/>");
        }
        return stream.ToArray();
    }

    private static void Write(OpenXmlPart part, string xml)
    {
        using var partStream = part.GetStream(FileMode.Create);
        var bytes = Encoding.UTF8.GetBytes(xml);
        partStream.Write(bytes, 0, bytes.Length);
    }

    /// <summary>The <c>margin-bottom</c> of each paragraph whose text starts with "Item", in document order.</summary>
    private static string[] ItemSpacingAfter(byte[] docx)
    {
        var html = XElement.Parse(HtmlConversionOps.ConvertToHtml(docx, new HtmlConversionOptions { FabricateCssClasses = false }));
        return html.Descendants()
            .Where(e => e.Name.LocalName is "p" or "li" && e.Value.Contains("Item"))
            .Select(e => ((string?)e.Attribute("style") ?? string.Empty)
                .Split(';')
                .Select(declaration => declaration.Trim())
                .FirstOrDefault(declaration => declaration.StartsWith("margin-bottom:"))
                ?.Substring("margin-bottom:".Length).Trim() ?? "(none)")
            .ToArray();
    }

    [Fact]
    public void ConsecutiveNumberedParagraphs_KeepTheirSpacingAfter()
    {
        var spacing = ItemSpacingAfter(Docx(Item("Item one") + Item("Item two") + Item("Item three"), contextualSpacing: false));

        Assert.Equal(new[] { "12pt", "12pt", "12pt" }, spacing);
    }

    [Fact]
    public void ContextualSpacing_StillRemovesSpacingBetweenSameStyleListItems()
    {
        // The style carries w:contextualSpacing, as Word's own List Paragraph style does.
        var spacing = ItemSpacingAfter(Docx(Item("Item one") + Item("Item two") + Item("Item three"), contextualSpacing: true));

        Assert.Equal(new[] { "0", "0", "12pt" }, spacing);
    }

    [Fact]
    public void ContextualSpacingTurnedOff_KeepsTheSpacing()
    {
        // w:val="0" turns the property off; it does not remove spacing.
        var spacing = ItemSpacingAfter(Docx(Item("Item one") + Item("Item two") + Item("Item three"),
            contextualSpacing: true, contextual: "<w:contextualSpacing w:val=\"0\"/>"));

        Assert.Equal(new[] { "12pt", "12pt", "12pt" }, spacing);
    }
}
