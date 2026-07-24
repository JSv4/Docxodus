#nullable enable

using Docxodus.Ir;
using Xunit;

namespace Docxodus.Tests.Ir;

public class IrFormatTests
{
    [Fact]
    public void IrRunFormat_RecordEquality_IncludesUnmodeledDigest()
    {
        var digestA = IrHash.Compute("unmodeled-a");
        var digestB = IrHash.Compute("unmodeled-b");

        var run1 = new IrRunFormat
        {
            Bold = true,
            Italic = false,
            Underline = new IrUnderline(IrUnderlineKind.Single, "FF0000"),
            SizeHalfPoints = 24,
            ColorHex = "auto",
            Shading = new IrShading("<w:shd w:fill=\"FFFF00\"/>"),
            UnmodeledDigest = digestA,
        };
        var run2 = new IrRunFormat
        {
            Bold = true,
            Italic = false,
            Underline = new IrUnderline(IrUnderlineKind.Single, "FF0000"),
            SizeHalfPoints = 24,
            ColorHex = "auto",
            Shading = new IrShading("<w:shd w:fill=\"FFFF00\"/>"),
            UnmodeledDigest = digestA,
        };
        var run3 = run2 with { UnmodeledDigest = digestB };
        var run4 = run2 with { Shading = new IrShading("<w:shd w:fill=\"00FF00\"/>") };

        Assert.Equal(run1, run2);
        Assert.Equal(run1.GetHashCode(), run2.GetHashCode());

        Assert.NotEqual(run1, run3);
        Assert.NotEqual(run1, run4);
    }

    [Fact]
    public void IrSectionFormat_RecordEquality_IncludesUnmodeledDigest()
    {
        var digestA = IrHash.Compute("section-unmodeled-a");
        var digestB = IrHash.Compute("section-unmodeled-b");

        var section1 = new IrSectionFormat
        {
            PageWidthTwips = 12240,
            PageHeightTwips = 15840,
            Landscape = false,
            MarginTopTwips = 1440,
            MarginBottomTwips = 1440,
            SectionType = "nextPage",
            UnmodeledDigest = digestA,
        };
        var section2 = new IrSectionFormat
        {
            PageWidthTwips = 12240,
            PageHeightTwips = 15840,
            Landscape = false,
            MarginTopTwips = 1440,
            MarginBottomTwips = 1440,
            SectionType = "nextPage",
            UnmodeledDigest = digestA,
        };
        var section3 = section2 with { UnmodeledDigest = digestB };

        Assert.Equal(section1, section2);
        Assert.Equal(section1.GetHashCode(), section2.GetHashCode());

        Assert.NotEqual(section1, section3);
    }

    [Fact]
    public void IrParaFormat_RecordEquality()
    {
        var digest = IrHash.Compute("para-unmodeled");

        var para1 = new IrParaFormat
        {
            StyleId = "Heading1",
            Justification = IrJustification.Both,
            IndentFirstLineTwips = -360,
            LineSpacing = new IrLineSpacing(240, IrLineSpacingRule.Auto),
            OutlineLevel = 0,
            KeepNext = true,
            UnmodeledDigest = digest,
        };
        var para2 = new IrParaFormat
        {
            StyleId = "Heading1",
            Justification = IrJustification.Both,
            IndentFirstLineTwips = -360,
            LineSpacing = new IrLineSpacing(240, IrLineSpacingRule.Auto),
            OutlineLevel = 0,
            KeepNext = true,
            UnmodeledDigest = digest,
        };
        var para3 = para2 with { LineSpacing = new IrLineSpacing(360, IrLineSpacingRule.Exact) };

        Assert.Equal(para1, para2);
        Assert.Equal(para1.GetHashCode(), para2.GetHashCode());

        Assert.NotEqual(para1, para3);
    }

    [Fact]
    public void IrListInfo_Equality()
    {
        var list1 = new IrListInfo(NumId: 3, AbstractNumId: 7, Ilvl: 1,
                                   NumberFormat: "decimal", StartOverride: null, FromStyle: false);
        var list2 = new IrListInfo(NumId: 3, AbstractNumId: 7, Ilvl: 1,
                                   NumberFormat: "decimal", StartOverride: null, FromStyle: false);
        var list3 = list2 with { Ilvl = 2 };

        Assert.Equal(list1, list2);
        Assert.Equal(list1.GetHashCode(), list2.GetHashCode());

        Assert.NotEqual(list1, list3);
    }
}
