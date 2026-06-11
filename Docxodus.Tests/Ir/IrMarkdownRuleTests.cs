#nullable enable

using Docxodus;
using Docxodus.Ir;
using Xunit;

namespace Docxodus.Tests.Ir;

/// <summary>
/// Per-rule pins for the IR markdown emitter (M1.4 Task 1). Each test builds a tiny DOCX exercising
/// one ported emission rule and asserts the IR path's markdown is byte-equal to the oracle's, so the
/// rule stays equivalent even when no corpus fixture exercises it. Default settings throughout.
/// </summary>
public class IrMarkdownRuleTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private static void AssertEquivalent(WmlDocument doc)
    {
        var settings = new WmlToMarkdownConverterSettings();
        // The oracle mutates bytes (persists Unids) — give it its own copy.
        var oracle = WmlToMarkdownConverter.Convert(new WmlDocument(doc), settings);
        var ir = IrMarkdownEmitter.Emit(IrReader.Read(new WmlDocument(doc)), settings);
        Assert.Equal(oracle.Markdown, ir.Markdown);
    }

    [Fact]
    public void Rule_PlainParagraph()
    {
        AssertEquivalent(IrTestDocuments.Create("Hello world.", "Second paragraph."));
    }

    [Fact]
    public void Rule_EmptyParagraph_AnchorOnly()
    {
        // Default EmptyParagraphMode.AnchorOnly: a runless paragraph emits the anchor with the
        // dangling separator space trimmed.
        AssertEquivalent(IrTestDocuments.FromBodyXml(
            "<w:p/><w:p><w:r><w:t>after</w:t></w:r></w:p>"));
    }

    [Theory]
    [InlineData("Heading1", 1)]
    [InlineData("Heading2", 2)]
    [InlineData("Heading3", 3)]
    [InlineData("Heading7", 7)]
    [InlineData("Title", 1)]
    [InlineData("Subtitle", 2)]
    public void Rule_HeadingLevels(string styleId, int _)
    {
        var body =
            $"<w:p><w:pPr><w:pStyle w:val=\"{styleId}\"/></w:pPr>" +
            "<w:r><w:t>The Heading</w:t></w:r></w:p>";
        var styles =
            $"<w:style w:type=\"paragraph\" w:styleId=\"{styleId}\"><w:name w:val=\"{styleId}\"/></w:style>";
        AssertEquivalent(IrTestDocuments.FromBodyAndStylesXml(body, styles));
    }

    [Fact]
    public void Rule_Bold()
    {
        AssertEquivalent(IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r></w:p>"));
    }

    [Fact]
    public void Rule_Italic()
    {
        AssertEquivalent(IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r></w:p>"));
    }

    [Fact]
    public void Rule_BoldItalic_Merged()
    {
        AssertEquivalent(IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>both</w:t></w:r>" +
            "<w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>more</w:t></w:r></w:p>"));
    }

    [Fact]
    public void Rule_Strike()
    {
        AssertEquivalent(IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:rPr><w:strike/></w:rPr><w:t>gone</w:t></w:r></w:p>"));
    }

    [Fact]
    public void Rule_ToggleOffWithVal0_IsNotBold()
    {
        // w:b w:val="0" is an explicit toggle-off → no ** delimiters.
        AssertEquivalent(IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:rPr><w:b w:val=\"0\"/></w:rPr><w:t>plain</w:t></w:r></w:p>"));
    }

    [Fact]
    public void Rule_EscapingMarkdownMetacharacters()
    {
        // Every markdown metachar must be backslash-escaped identically in both paths.
        AssertEquivalent(IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t xml:space=\"preserve\">a*b_c`d#e+f-g!h|i&gt;j~k[l](m){n}o</w:t></w:r></w:p>"));
    }

    [Fact]
    public void Rule_Hyperlink_External()
    {
        // The builder declares xmlns:w only, so declare xmlns:r on the hyperlink element itself.
        var body =
            $"<w:p><w:hyperlink xmlns:r=\"{R}\" r:id=\"rId99\"><w:r><w:t>click here</w:t></w:r></w:hyperlink></w:p>";
        AssertEquivalent(IrTestDocuments.FromBodyXmlWithHyperlinks(body, ("rId99", "https://example.com/")));
    }

    [Fact]
    public void Rule_Hyperlink_InternalAnchor()
    {
        var body =
            "<w:p><w:hyperlink w:anchor=\"Bookmark1\"><w:r><w:t>jump</w:t></w:r></w:hyperlink></w:p>";
        AssertEquivalent(IrTestDocuments.FromBodyXml(body));
    }

    [Fact]
    public void Rule_Tab()
    {
        AssertEquivalent(IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r></w:p>"));
    }

    [Fact]
    public void Rule_LineBreak()
    {
        AssertEquivalent(IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>a</w:t><w:br/><w:t>b</w:t></w:r></w:p>"));
    }

    [Fact]
    public void Rule_BulletList_RendersDash()
    {
        // A bullet-format numbering definition. Both paths render "-" for bullet levels.
        var body =
            "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr>" +
            "<w:r><w:t>first</w:t></w:r></w:p>" +
            "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr>" +
            "<w:r><w:t>second</w:t></w:r></w:p>";
        var numbering =
            "<w:abstractNum w:abstractNumId=\"0\">" +
            "<w:lvl w:ilvl=\"0\"><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"·\"/></w:lvl>" +
            "</w:abstractNum>" +
            "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>";
        AssertEquivalent(IrTestDocuments.FromParts(body, stylesInnerXml: "", numberingInnerXml: numbering));
    }

    [Fact]
    public void Rule_NestedBulletList_SymbolGlyphs_Indentation()
    {
        // Both levels use a NON-alphanumeric bullet glyph, which the oracle's ResolveListMarker
        // collapses to "-" (its rule: a single non-letter-or-digit resolved marker → "-"). Indent
        // is 2 spaces per ilvl. A level whose lvlText is an alphanumeric glyph (e.g. "o") is NOT
        // collapsed by the oracle and needs the IR counter walk — TODO(M1.4-T3), off the must-pass
        // list — so this pin deliberately uses symbol glyphs only.
        var body =
            "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr>" +
            "<w:r><w:t>top</w:t></w:r></w:p>" +
            "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"1\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr>" +
            "<w:r><w:t>nested</w:t></w:r></w:p>";
        var numbering =
            "<w:abstractNum w:abstractNumId=\"0\">" +
            "<w:lvl w:ilvl=\"0\"><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"·\"/></w:lvl>" +
            "<w:lvl w:ilvl=\"1\"><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"§\"/></w:lvl>" +
            "</w:abstractNum>" +
            "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>";
        AssertEquivalent(IrTestDocuments.FromParts(body, stylesInnerXml: "", numberingInnerXml: numbering));
    }

    [Fact]
    public void Rule_CodeRun_MonospaceFont()
    {
        // A Consolas run is treated as a `code` span by both paths.
        AssertEquivalent(IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:rPr><w:rFonts w:ascii=\"Consolas\"/></w:rPr><w:t>x = 1</w:t></w:r></w:p>"));
    }
}
