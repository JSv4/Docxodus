#nullable enable

using System.Xml.Linq;
using Docxodus;
using Docxodus.Ir;
using Xunit;

namespace Docxodus.Tests.Ir;

/// <summary>
/// Facts about the M1.3 style / numbering / theme registries populated by
/// <see cref="IrReader"/>. Each test builds a minimal programmatic package with just the parts it
/// needs and asserts the resolved registry shape. Totality cases (missing/malformed parts → Empty
/// registries) guard the corpus-totality contract: registry building now runs for every file.
/// </summary>
public class IrRegistryTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // --- styles -----------------------------------------------------------

    [Fact]
    public void StyleRegistry_PopulatesIdNameBasedOnTypeDefaultAndPPr()
    {
        var styles =
            "<w:docDefaults>" +
            "  <w:rPrDefault><w:rPr><w:sz w:val=\"20\"/></w:rPr></w:rPrDefault>" +
            "  <w:pPrDefault><w:pPr><w:spacing w:after=\"120\"/></w:pPr></w:pPrDefault>" +
            "</w:docDefaults>" +
            "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\">" +
            "  <w:name w:val=\"Normal\"/>" +
            "  <w:pPr><w:jc w:val=\"both\"/></w:pPr>" +
            "  <w:rPr><w:b/></w:rPr>" +
            "</w:style>" +
            "<w:style w:type=\"paragraph\" w:styleId=\"Heading1\">" +
            "  <w:name w:val=\"heading 1\"/>" +
            "  <w:basedOn w:val=\"Normal\"/>" +
            "</w:style>";

        var ir = IrReader.Read(IrTestDocuments.FromParts("<w:p/>", stylesInnerXml: styles));
        var reg = ir.Styles;

        Assert.Equal(2, reg.Styles.Count);
        Assert.Equal("Normal", reg.DefaultParagraphStyleId);

        var normal = reg.Styles["Normal"];
        Assert.Equal("Normal", normal.Id);
        Assert.Equal("Normal", normal.Name);
        Assert.Null(normal.BasedOn);
        Assert.Equal("paragraph", normal.Type);
        Assert.True(normal.IsDefault);
        Assert.NotNull(normal.PPr);
        Assert.NotNull(normal.RPr);

        var h1 = reg.Styles["Heading1"];
        Assert.Equal("heading 1", h1.Name);
        Assert.Equal("Normal", h1.BasedOn);
        Assert.False(h1.IsDefault);

        // docDefaults clones present.
        Assert.NotNull(reg.DocDefaultsPPr);
        Assert.NotNull(reg.DocDefaultsRPr);
        Assert.Equal("120", (string?)reg.DocDefaultsPPr!.Element(W + "spacing")?.Attribute(W + "after"));
    }

    [Fact]
    public void StyleRegistry_TypeDefaultsToParagraphWhenMissing()
    {
        var styles = "<w:style w:styleId=\"Bare\"><w:name w:val=\"Bare\"/></w:style>";
        var ir = IrReader.Read(IrTestDocuments.FromParts("<w:p/>", stylesInnerXml: styles));
        Assert.Equal("paragraph", ir.Styles.Styles["Bare"].Type);
    }

    [Fact]
    public void StyleRegistry_StyleWithoutStyleId_IsSkipped()
    {
        var styles =
            "<w:style w:type=\"paragraph\"><w:name w:val=\"NoId\"/></w:style>" +
            "<w:style w:type=\"paragraph\" w:styleId=\"Good\"><w:name w:val=\"Good\"/></w:style>";
        var ir = IrReader.Read(IrTestDocuments.FromParts("<w:p/>", stylesInnerXml: styles));

        Assert.Single(ir.Styles.Styles);
        Assert.True(ir.Styles.Styles.ContainsKey("Good"));
    }

    [Fact]
    public void StyleRegistry_DuplicateStyleId_FirstWins()
    {
        var styles =
            "<w:style w:type=\"paragraph\" w:styleId=\"Dup\"><w:name w:val=\"First\"/></w:style>" +
            "<w:style w:type=\"paragraph\" w:styleId=\"Dup\"><w:name w:val=\"Second\"/></w:style>";
        var ir = IrReader.Read(IrTestDocuments.FromParts("<w:p/>", stylesInnerXml: styles));

        Assert.Single(ir.Styles.Styles);
        Assert.Equal("First", ir.Styles.Styles["Dup"].Name);
    }

    // --- numbering --------------------------------------------------------

    [Fact]
    public void NumberingRegistry_PopulatesAbstractNumLevelsAndNums()
    {
        var numbering =
            "<w:abstractNum w:abstractNumId=\"0\">" +
            "  <w:lvl w:ilvl=\"0\">" +
            "    <w:start w:val=\"1\"/>" +
            "    <w:numFmt w:val=\"bullet\"/>" +
            "    <w:lvlText w:val=\"•\"/>" +
            "    <w:pPr><w:ind w:left=\"720\"/></w:pPr>" +
            "  </w:lvl>" +
            "  <w:lvl w:ilvl=\"1\">" +
            "    <w:numFmt w:val=\"decimal\"/>" +
            "    <w:lvlText w:val=\"%2.\"/>" +
            "  </w:lvl>" +
            "</w:abstractNum>" +
            "<w:num w:numId=\"3\">" +
            "  <w:abstractNumId w:val=\"0\"/>" +
            "  <w:lvlOverride w:ilvl=\"0\"><w:startOverride w:val=\"5\"/></w:lvlOverride>" +
            "</w:num>";

        var ir = IrReader.Read(IrTestDocuments.FromParts("<w:p/>", numberingInnerXml: numbering));
        var reg = ir.Numbering;

        Assert.Single(reg.AbstractNums);
        var abs = reg.AbstractNums[0];
        Assert.Equal(2, abs.Levels.Count);

        var lvl0 = abs.Levels[0];
        Assert.Equal(0, lvl0.Ilvl);
        Assert.Equal("bullet", lvl0.NumberFormat);
        Assert.Equal(1, lvl0.Start);
        Assert.Equal("•", lvl0.LvlText);
        Assert.NotNull(lvl0.PPr);

        var lvl1 = abs.Levels[1];
        Assert.Equal("decimal", lvl1.NumberFormat);
        Assert.Null(lvl1.Start);
        Assert.Equal("%2.", lvl1.LvlText);
        Assert.Null(lvl1.PPr);

        Assert.Single(reg.Nums);
        var num = reg.Nums[3];
        Assert.Equal(3, num.NumId);
        Assert.Equal(0, num.AbstractNumId);
        Assert.Equal(5, num.StartOverrides[0]);
    }

    [Fact]
    public void NumberingRegistry_NumFmtDefaultsToDecimal()
    {
        var numbering =
            "<w:abstractNum w:abstractNumId=\"0\"><w:lvl w:ilvl=\"0\"/></w:abstractNum>";
        var ir = IrReader.Read(IrTestDocuments.FromParts("<w:p/>", numberingInnerXml: numbering));
        Assert.Equal("decimal", ir.Numbering.AbstractNums[0].Levels[0].NumberFormat);
    }

    [Fact]
    public void NumberingRegistry_NumStyleLinkAbstractNum_LandsWithNoLevels()
    {
        // An abstractNum that borrows its levels via numStyleLink carries no explicit w:lvl.
        // We record it as-is (empty levels) rather than chasing the indirection (M1.4+ TODO).
        var numbering =
            "<w:abstractNum w:abstractNumId=\"7\">" +
            "  <w:numStyleLink w:val=\"MyListStyle\"/>" +
            "</w:abstractNum>";
        var ir = IrReader.Read(IrTestDocuments.FromParts("<w:p/>", numberingInnerXml: numbering));

        Assert.Empty(ir.Numbering.AbstractNums[7].Levels);
    }

    [Fact]
    public void NumberingRegistry_DuplicateAbstractNumId_FirstWins()
    {
        var numbering =
            "<w:abstractNum w:abstractNumId=\"0\"><w:lvl w:ilvl=\"0\"><w:numFmt w:val=\"bullet\"/></w:lvl></w:abstractNum>" +
            "<w:abstractNum w:abstractNumId=\"0\"><w:lvl w:ilvl=\"0\"><w:numFmt w:val=\"decimal\"/></w:lvl></w:abstractNum>";
        var ir = IrReader.Read(IrTestDocuments.FromParts("<w:p/>", numberingInnerXml: numbering));

        Assert.Single(ir.Numbering.AbstractNums);
        Assert.Equal("bullet", ir.Numbering.AbstractNums[0].Levels[0].NumberFormat);
    }

    // --- theme ------------------------------------------------------------

    [Fact]
    public void ThemeFonts_PopulatesMajorAndMinorAscii()
    {
        var fontScheme =
            "<a:majorFont><a:latin typeface=\"Calibri Light\"/></a:majorFont>" +
            "<a:minorFont><a:latin typeface=\"Calibri\"/></a:minorFont>";
        var ir = IrReader.Read(IrTestDocuments.FromParts("<w:p/>", themeFontSchemeInnerXml: fontScheme));

        Assert.Equal("Calibri Light", ir.ThemeFonts.MajorAscii);
        Assert.Equal("Calibri", ir.ThemeFonts.MinorAscii);
    }

    // --- totality (missing / malformed parts → Empty) ---------------------

    [Fact]
    public void MissingParts_YieldEmptyRegistries()
    {
        // FromBodyXml has a StyleDefinitionsPart but no styles, no numbering, no theme.
        var ir = IrReader.Read(IrTestDocuments.FromBodyXml("<w:p/>"));

        Assert.Empty(ir.Styles.Styles);
        Assert.Null(ir.Styles.DefaultParagraphStyleId);
        Assert.Same(IrNumberingRegistry.Empty, ir.Numbering);
        Assert.Same(IrThemeFonts.Empty, ir.ThemeFonts);
    }

    [Fact]
    public void EmptyRegistries_AreReferenceEqualToEmptyMembers()
    {
        Assert.Same(IrStyleRegistry.Empty, IrStyleRegistry.Empty);
        Assert.Empty(IrStyleRegistry.Empty.Styles);
        Assert.Empty(IrNumberingRegistry.Empty.Nums);
        Assert.Empty(IrNumberingRegistry.Empty.AbstractNums);
        Assert.Null(IrThemeFonts.Empty.MajorAscii);
        Assert.Null(IrThemeFonts.Empty.MinorAscii);
    }
}
