#nullable enable

using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Ir;
using Xunit;

namespace Docxodus.Tests.Ir;

/// <summary>
/// Hand-computed cascade expectations for <see cref="IrEffectiveFormats"/> (§5.2): docDefaults →
/// style chain (basedOn, root-first) → direct, with per-field non-null-wins merge, theme-font
/// indirection, cycle guarding, the default paragraph style, and a cache-consistency check. One
/// case cross-checks against <see cref="FormattingAssembler"/> on a real assembled document.
/// </summary>
public class IrEffectiveFormatTests
{
    // Read a document's first body paragraph (and its first text run) for resolution.
    private static (IrEffectiveFormats Eff, IrParagraph Para) FirstParagraph(WmlDocument doc)
    {
        var ir = IrReader.Read(doc);
        var para = ir.Body.Blocks.OfType<IrParagraph>().First();
        return (new IrEffectiveFormats(ir), para);
    }

    private static IrRunFormat FirstRunFormat(IrParagraph p) =>
        p.Inlines.OfType<IrTextRun>().First().Format;

    // --- paragraph cascade ------------------------------------------------

    [Fact]
    public void Paragraph_DocDefaults_StyleChain_AndDirect_AllContribute()
    {
        // docDefaults size 20; style Base bold + size 24; style Derived basedOn Base + italic;
        // paragraph styled Derived with a direct size 28. Expected effective run/para facts:
        //   Bold=true (from Base), Italic=true (from Derived), SizeHalfPoints=28 (direct override).
        var styles =
            "<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val=\"20\"/></w:rPr></w:rPrDefault></w:docDefaults>" +
            "<w:style w:type=\"paragraph\" w:styleId=\"Base\">" +
            "  <w:rPr><w:b/><w:sz w:val=\"24\"/></w:rPr>" +
            "</w:style>" +
            "<w:style w:type=\"paragraph\" w:styleId=\"Derived\">" +
            "  <w:basedOn w:val=\"Base\"/>" +
            "  <w:rPr><w:i/></w:rPr>" +
            "</w:style>";
        var body =
            "<w:p><w:pPr><w:pStyle w:val=\"Derived\"/></w:pPr>" +
            "<w:r><w:rPr><w:sz w:val=\"28\"/></w:rPr><w:t>x</w:t></w:r></w:p>";

        var (eff, para) = FirstParagraph(IrTestDocuments.FromParts(body, stylesInnerXml: styles));
        var run = eff.ResolveRun(para, FirstRunFormat(para));

        Assert.True(run.Bold);
        Assert.True(run.Italic);
        Assert.Equal(28, run.SizeHalfPoints);
    }

    [Fact]
    public void Paragraph_NullFieldInDerived_InheritsFromBase()
    {
        // Base sets jc=center; Derived sets nothing on jc → effective justification is center.
        var styles =
            "<w:style w:type=\"paragraph\" w:styleId=\"Base\">" +
            "  <w:pPr><w:jc w:val=\"center\"/></w:pPr>" +
            "</w:style>" +
            "<w:style w:type=\"paragraph\" w:styleId=\"Derived\">" +
            "  <w:basedOn w:val=\"Base\"/>" +
            "  <w:pPr><w:keepNext/></w:pPr>" +
            "</w:style>";
        var body = "<w:p><w:pPr><w:pStyle w:val=\"Derived\"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>";

        var (eff, para) = FirstParagraph(IrTestDocuments.FromParts(body, stylesInnerXml: styles));
        var fmt = eff.ResolveParagraph(para);

        Assert.Equal(IrJustification.Center, fmt.Justification);
        Assert.True(fmt.KeepNext);
    }

    [Fact]
    public void Paragraph_DerivedOverridesBase_RootFirstApplication()
    {
        // Base jc=left, Derived jc=right → derived wins (root-first means derived is the topmost
        // style layer below direct).
        var styles =
            "<w:style w:type=\"paragraph\" w:styleId=\"Base\"><w:pPr><w:jc w:val=\"left\"/></w:pPr></w:style>" +
            "<w:style w:type=\"paragraph\" w:styleId=\"Derived\">" +
            "  <w:basedOn w:val=\"Base\"/><w:pPr><w:jc w:val=\"right\"/></w:pPr></w:style>";
        var body = "<w:p><w:pPr><w:pStyle w:val=\"Derived\"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>";

        var (eff, para) = FirstParagraph(IrTestDocuments.FromParts(body, stylesInnerXml: styles));
        Assert.Equal(IrJustification.Right, eff.ResolveParagraph(para).Justification);
    }

    [Fact]
    public void Paragraph_DefaultStyleApplies_WhenNoPStyle()
    {
        // The default paragraph style sets jc=center; a paragraph with NO pStyle inherits it.
        var styles =
            "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\">" +
            "  <w:pPr><w:jc w:val=\"center\"/></w:pPr>" +
            "</w:style>";
        var body = "<w:p><w:r><w:t>x</w:t></w:r></w:p>";

        var (eff, para) = FirstParagraph(IrTestDocuments.FromParts(body, stylesInnerXml: styles));
        Assert.Equal(IrJustification.Center, eff.ResolveParagraph(para).Justification);
    }

    // --- run cascade ------------------------------------------------------

    [Fact]
    public void Run_ParagraphStyleColor_PlusDirectBold()
    {
        // Paragraph style rPr color red; run direct bold → effective red + bold.
        var styles =
            "<w:style w:type=\"paragraph\" w:styleId=\"Body\">" +
            "  <w:rPr><w:color w:val=\"FF0000\"/></w:rPr>" +
            "</w:style>";
        var body =
            "<w:p><w:pPr><w:pStyle w:val=\"Body\"/></w:pPr>" +
            "<w:r><w:rPr><w:b/></w:rPr><w:t>x</w:t></w:r></w:p>";

        var (eff, para) = FirstParagraph(IrTestDocuments.FromParts(body, stylesInnerXml: styles));
        var run = eff.ResolveRun(para, FirstRunFormat(para));

        Assert.Equal("FF0000", run.ColorHex);
        Assert.True(run.Bold);
    }

    [Fact]
    public void Run_CharacterStyleOverridesParagraphStyleColor()
    {
        // Paragraph style color red; a character style (named by the run's rStyle) color blue →
        // the character style wins (it sits above the paragraph-style chain).
        var styles =
            "<w:style w:type=\"paragraph\" w:styleId=\"Body\">" +
            "  <w:rPr><w:color w:val=\"FF0000\"/></w:rPr>" +
            "</w:style>" +
            "<w:style w:type=\"character\" w:styleId=\"Emph\">" +
            "  <w:rPr><w:color w:val=\"0000FF\"/></w:rPr>" +
            "</w:style>";
        var body =
            "<w:p><w:pPr><w:pStyle w:val=\"Body\"/></w:pPr>" +
            "<w:r><w:rPr><w:rStyle w:val=\"Emph\"/></w:rPr><w:t>x</w:t></w:r></w:p>";

        var (eff, para) = FirstParagraph(IrTestDocuments.FromParts(body, stylesInnerXml: styles));
        var run = eff.ResolveRun(para, FirstRunFormat(para));

        Assert.Equal("0000FF", run.ColorHex);
    }

    // --- theme fonts ------------------------------------------------------

    [Fact]
    public void Run_ThemeFont_ResolvesAsciiThemeToMinorAscii()
    {
        // docDefaults rPr carries asciiTheme=minorHAnsi (the realistic location for the theme
        // indirection); theme MinorAscii="Calibri" → effective FontAscii Calibri. Theme resolution
        // happens at the rPr-bearing layer the resolver maps; the direct run carries no font.
        var theme = "<a:minorFont><a:latin typeface=\"Calibri\"/></a:minorFont>";
        var styles =
            "<w:docDefaults><w:rPrDefault><w:rPr>" +
            "<w:rFonts w:asciiTheme=\"minorHAnsi\"/></w:rPr></w:rPrDefault></w:docDefaults>";
        var body = "<w:p><w:r><w:t>x</w:t></w:r></w:p>";

        var (eff, para) = FirstParagraph(
            IrTestDocuments.FromParts(body, stylesInnerXml: styles, themeFontSchemeInnerXml: theme));
        var run = eff.ResolveRun(para, FirstRunFormat(para));

        Assert.Equal("Calibri", run.FontAscii);
    }

    [Fact]
    public void Run_DirectAsciiBeatsThemeFont()
    {
        // docDefaults resolves to the theme face Calibri; the run's direct @w:ascii=Arial is a
        // higher layer and wins (later non-null field wins).
        var theme = "<a:minorFont><a:latin typeface=\"Calibri\"/></a:minorFont>";
        var styles =
            "<w:docDefaults><w:rPrDefault><w:rPr>" +
            "<w:rFonts w:asciiTheme=\"minorHAnsi\"/></w:rPr></w:rPrDefault></w:docDefaults>";
        var body =
            "<w:p><w:r><w:rPr><w:rFonts w:ascii=\"Arial\"/></w:rPr><w:t>x</w:t></w:r></w:p>";

        var (eff, para) = FirstParagraph(
            IrTestDocuments.FromParts(body, stylesInnerXml: styles, themeFontSchemeInnerXml: theme));
        var run = eff.ResolveRun(para, FirstRunFormat(para));

        Assert.Equal("Arial", run.FontAscii);
    }

    // --- cycle guard ------------------------------------------------------

    [Fact]
    public void StyleBasedOnItself_NoHang_DirectFactsStillResolve()
    {
        // A style whose basedOn points at itself must not hang the chain walk; direct facts still
        // resolve, and the style's own props apply once.
        var styles =
            "<w:style w:type=\"paragraph\" w:styleId=\"Loop\">" +
            "  <w:basedOn w:val=\"Loop\"/>" +
            "  <w:pPr><w:jc w:val=\"center\"/></w:pPr>" +
            "  <w:rPr><w:b/></w:rPr>" +
            "</w:style>";
        var body =
            "<w:p><w:pPr><w:pStyle w:val=\"Loop\"/></w:pPr>" +
            "<w:r><w:rPr><w:i/></w:rPr><w:t>x</w:t></w:r></w:p>";

        var (eff, para) = FirstParagraph(IrTestDocuments.FromParts(body, stylesInnerXml: styles));

        var pf = eff.ResolveParagraph(para);
        Assert.Equal(IrJustification.Center, pf.Justification);

        var run = eff.ResolveRun(para, FirstRunFormat(para));
        Assert.True(run.Bold);   // from the (self-based) style
        Assert.True(run.Italic); // direct
    }

    // --- cache consistency ------------------------------------------------

    [Fact]
    public void Cache_SameStyle_TwoParagraphs_ConsistentResults()
    {
        // Two paragraphs of the same style resolve to the same effective facts: the second hits the
        // memoized style-chain contribution and must agree with the first.
        var styles =
            "<w:style w:type=\"paragraph\" w:styleId=\"Body\">" +
            "  <w:pPr><w:jc w:val=\"center\"/></w:pPr>" +
            "  <w:rPr><w:b/><w:sz w:val=\"24\"/></w:rPr>" +
            "</w:style>";
        var body =
            "<w:p><w:pPr><w:pStyle w:val=\"Body\"/></w:pPr><w:r><w:t>a</w:t></w:r></w:p>" +
            "<w:p><w:pPr><w:pStyle w:val=\"Body\"/></w:pPr><w:r><w:t>b</w:t></w:r></w:p>";

        var ir = IrReader.Read(IrTestDocuments.FromParts(body, stylesInnerXml: styles));
        var eff = new IrEffectiveFormats(ir);
        var paras = ir.Body.Blocks.OfType<IrParagraph>().ToList();

        var p0 = eff.ResolveParagraph(paras[0]);
        var p1 = eff.ResolveParagraph(paras[1]);
        Assert.Equal(p0.Justification, p1.Justification);
        Assert.Equal(IrJustification.Center, p1.Justification);

        var r0 = eff.ResolveRun(paras[0], FirstRunFormat(paras[0]));
        var r1 = eff.ResolveRun(paras[1], FirstRunFormat(paras[1]));
        Assert.Equal(r0.Bold, r1.Bold);
        Assert.Equal(r0.SizeHalfPoints, r1.SizeHalfPoints);
        Assert.True(r1.Bold);
        Assert.Equal(24, r1.SizeHalfPoints);
    }

    // --- FormattingAssembler cross-check ----------------------------------

    [Fact]
    public void CrossCheck_FormattingAssembler_BoldAndSize_MatchResolveRun()
    {
        // Build a real document with a docDefault size, a Base style (bold+size), a Derived style
        // (basedOn Base, italic), and a paragraph styled Derived with a direct size override. Run
        // FormattingAssembler on a copy — it rolls the fully-resolved rPr onto each run — and compare
        // its w:b / w:sz on the run against IrEffectiveFormats.ResolveRun on the same source.
        var styles =
            "<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val=\"20\"/></w:rPr></w:rPrDefault></w:docDefaults>" +
            "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>" +
            "<w:style w:type=\"paragraph\" w:styleId=\"Base\">" +
            "  <w:name w:val=\"Base\"/><w:rPr><w:b/><w:sz w:val=\"24\"/></w:rPr></w:style>" +
            "<w:style w:type=\"paragraph\" w:styleId=\"Derived\">" +
            "  <w:name w:val=\"Derived\"/><w:basedOn w:val=\"Base\"/><w:rPr><w:i/></w:rPr></w:style>";
        var body =
            "<w:p><w:pPr><w:pStyle w:val=\"Derived\"/></w:pPr>" +
            "<w:r><w:rPr><w:sz w:val=\"28\"/></w:rPr><w:t>x</w:t></w:r></w:p>";
        var doc = IrTestDocuments.FromParts(body, stylesInnerXml: styles);

        // IR effective resolution.
        var (eff, para) = FirstParagraph(doc);
        var effRun = eff.ResolveRun(para, FirstRunFormat(para));

        // FormattingAssembler reference on a copy.
        var (faBold, faSize) = AssembledRunBoldAndSize(doc);

        Assert.Equal(faBold, effRun.Bold ?? false);
        Assert.Equal(faSize, effRun.SizeHalfPoints);
    }

    /// <summary>
    /// Run <see cref="FormattingAssembler"/> on a copy of <paramref name="doc"/> and read the
    /// first run's assembled bold flag and half-point size from its rolled-up <c>w:rPr</c>.
    /// </summary>
    private static (bool Bold, int? SizeHalfPoints) AssembledRunBoldAndSize(WmlDocument doc)
    {
        var settings = new FormattingAssemblerSettings();
        var assembled = FormattingAssembler.AssembleFormatting(new WmlDocument(doc), settings);

        using var stream = new OpenXmlMemoryStreamDocument(assembled);
        using var wdoc = stream.GetWordprocessingDocument();
        var w = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var root = wdoc.MainDocumentPart!.GetXDocument().Root!;
        var run = root.Descendants(w + "r").First(r => r.Element(w + "t") is not null);
        var rPr = run.Element(w + "rPr");

        var bEl = rPr?.Element(w + "b");
        bool bold = bEl is not null
            && (string?)bEl.Attribute(w + "val") is not ("0" or "false" or "off");

        var szVal = (string?)rPr?.Element(w + "sz")?.Attribute(w + "val");
        int? size = int.TryParse(szVal, out var s) ? s : null;

        return (bold, size);
    }
}
