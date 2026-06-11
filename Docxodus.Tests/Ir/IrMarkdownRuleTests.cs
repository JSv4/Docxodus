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

    private static void AssertEquivalent(WmlDocument doc) =>
        AssertEquivalent(doc, new WmlToMarkdownConverterSettings());

    private static void AssertEquivalent(WmlDocument doc, WmlToMarkdownConverterSettings settings)
    {
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

    // --- M1.4-T2 rules: tables, images, section breaks, settings modes -------------------------

    /// <summary>A simple 2x2 table with short cells renders as a GFM pipe table in both paths.</summary>
    [Fact]
    public void Rule_SimpleTable_RendersAsGfm()
    {
        AssertEquivalent(IrTestDocuments.FromBodyXml(SimpleTableXml));
    }

    /// <summary>A horizontally-merged cell (w:gridSpan val>1) disqualifies GFM; both paths emit the
    /// opaque <c>```table rows/cols</c> block.</summary>
    [Fact]
    public void Rule_MergedCellTable_RendersAsOpaque()
    {
        var body =
            "<w:tbl><w:tr>" +
            "<w:tc><w:tcPr><w:gridSpan w:val=\"2\"/></w:tcPr><w:p><w:r><w:t>wide</w:t></w:r></w:p></w:tc>" +
            "</w:tr><w:tr>" +
            "<w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>" +
            "<w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc>" +
            "</w:tr></w:tbl>";
        AssertEquivalent(IrTestDocuments.FromBodyXml(body));
    }

    /// <summary>A vertically-merged cell (w:vMerge) also forces the opaque table path.</summary>
    [Fact]
    public void Rule_VMergeTable_RendersAsOpaque()
    {
        var body =
            "<w:tbl>" +
            "<w:tr><w:tc><w:tcPr><w:vMerge w:val=\"restart\"/></w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>" +
            "<w:tc><w:p><w:r><w:t>y</w:t></w:r></w:p></w:tc></w:tr>" +
            "<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>" +
            "<w:tc><w:p><w:r><w:t>z</w:t></w:r></w:p></w:tc></w:tr>" +
            "</w:tbl>";
        AssertEquivalent(IrTestDocuments.FromBodyXml(body));
    }

    /// <summary>A cell whose text exceeds <c>TableInlineCellMax</c> downgrades the table to opaque.
    /// Drive it with a low cap so a short fixture exercises the boundary in both paths.</summary>
    [Fact]
    public void Rule_OverLongCell_RendersAsOpaque_AtLowCap()
    {
        var settings = new WmlToMarkdownConverterSettings { TableInlineCellMax = 3 };
        var body =
            "<w:tbl><w:tr>" +
            "<w:tc><w:p><w:r><w:t>toolongcell</w:t></w:r></w:p></w:tc>" +
            "<w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc>" +
            "</w:tr></w:tbl>";
        AssertEquivalent(IrTestDocuments.FromBodyXml(body), settings);
    }

    /// <summary>Pipe characters inside a GFM cell are escaped and newlines collapsed identically.</summary>
    [Fact]
    public void Rule_GfmCell_EscapesPipes()
    {
        var body =
            "<w:tbl><w:tr>" +
            "<w:tc><w:p><w:r><w:t>a|b</w:t></w:r></w:p></w:tc>" +
            "<w:tc><w:p><w:r><w:t xml:space=\"preserve\"> c </w:t></w:r></w:p></w:tc>" +
            "</w:tr></w:tbl>";
        AssertEquivalent(IrTestDocuments.FromBodyXml(body));
    }

    /// <summary>An inline image: the oracle emits no image markup (it has no w:drawing emission path),
    /// so the paragraph projects empty — the IR matches byte-for-byte.</summary>
    [Fact]
    public void Rule_InlineImage_NoMarkupEmitted()
    {
        var body =
            "<w:p><w:r><w:drawing><wp:inline><wp:extent cx=\"100\" cy=\"100\"/>" +
            "<wp:docPr id=\"1\" name=\"Pic\" descr=\"alt\"/>" +
            "<a:graphic><a:graphicData><pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
            "<pic:blipFill><a:blip r:embed=\"rIdImg\"/></pic:blipFill></pic:pic>" +
            "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";
        AssertEquivalent(
            IrTestDocuments.FromBodyXmlWithImageParts(body, ("rIdImg", IrTestDocuments.TinyPng)));
    }

    /// <summary>An in-pPr <c>w:sectPr</c> (section transition) projects the <c>{#sec:…}</c> anchor plus
    /// a <c>---</c> thematic break after the paragraph; the trailing top-level body sectPr is metadata
    /// and emits nothing. Body-only (no headers/footers) so it isolates the section-break rule.</summary>
    [Fact]
    public void Rule_InlineSectionBreak()
    {
        var body =
            "<w:p><w:pPr><w:sectPr><w:type w:val=\"nextPage\"/></w:sectPr></w:pPr>" +
            "<w:r><w:t>before the break</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>after the break</w:t></w:r></w:p>" +
            "<w:sectPr><w:type w:val=\"nextPage\"/></w:sectPr>";
        AssertEquivalent(IrTestDocuments.FromBodyXml(body));
    }

    /// <summary>The three <see cref="AnchorIdRendering"/> modes must render anchor tokens identically
    /// in both paths — same per-(kind,scope) AnchorIdMap construction order, so abbreviations and
    /// sequential ids match byte-for-byte.</summary>
    [Theory]
    [InlineData(AnchorIdRendering.FullUnid)]
    [InlineData(AnchorIdRendering.Abbreviated)]
    [InlineData(AnchorIdRendering.Sequential)]
    public void Rule_AnchorIdRendering_Modes(AnchorIdRendering rendering)
    {
        var settings = new WmlToMarkdownConverterSettings { AnchorIdRendering = rendering };
        // Several blocks across kinds (plain + heading + table) so each bucket has >1 member,
        // exercising the abbreviation uniqueness search and the sequential counter.
        var body =
            "<w:p><w:r><w:t>one</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>two</w:t></w:r></w:p>" +
            SimpleTableXml +
            "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr><w:r><w:t>head</w:t></w:r></w:p>";
        var styles =
            "<w:style w:type=\"paragraph\" w:styleId=\"Heading1\"><w:name w:val=\"Heading1\"/></w:style>";
        AssertEquivalent(IrTestDocuments.FromBodyAndStylesXml(body, styles), settings);
    }

    /// <summary>The three <see cref="EmptyParagraphMode"/> values render runless paragraphs identically
    /// in both paths (anchor-only / ∅-marked / suppressed-from-output-and-index).</summary>
    [Theory]
    [InlineData(EmptyParagraphMode.AnchorOnly)]
    [InlineData(EmptyParagraphMode.MarkedEmpty)]
    [InlineData(EmptyParagraphMode.Suppress)]
    public void Rule_EmptyParagraph_Modes(EmptyParagraphMode mode)
    {
        var settings = new WmlToMarkdownConverterSettings { EmptyParagraphs = mode };
        AssertEquivalent(IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>a</w:t></w:r></w:p><w:p/><w:p><w:r><w:t>b</w:t></w:r></w:p>"), settings);
    }

    private const string SimpleTableXml =
        "<w:tbl>" +
        "<w:tr><w:tc><w:p><w:r><w:t>h1</w:t></w:r></w:p></w:tc>" +
        "<w:tc><w:p><w:r><w:t>h2</w:t></w:r></w:p></w:tc></w:tr>" +
        "<w:tr><w:tc><w:p><w:r><w:t>r1</w:t></w:r></w:p></w:tc>" +
        "<w:tc><w:p><w:r><w:t>r2</w:t></w:r></w:p></w:tc></w:tr>" +
        "</w:tbl>";
}
