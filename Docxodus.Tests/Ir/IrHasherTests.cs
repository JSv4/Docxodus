#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Docxodus.Ir;
using Xunit;

namespace Docxodus.Tests.Ir;

public class IrHasherTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace Pt = "http://powertools.codeplex.com/2011";
    private static readonly XNamespace Wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly XNamespace Wp14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";

    // --- Canonicalization -------------------------------------------------

    [Fact]
    public void Canonicalize_AttributeOrder_Irrelevant()
    {
        var a = new XElement(W + "p",
            new XAttribute(W + "alpha", "1"),
            new XAttribute(W + "beta", "2"),
            new XAttribute(W + "gamma", "3"));
        var b = new XElement(W + "p",
            new XAttribute(W + "gamma", "3"),
            new XAttribute(W + "alpha", "1"),
            new XAttribute(W + "beta", "2"));

        Assert.Equal(IrHasher.Canonicalize(a), IrHasher.Canonicalize(b));
        Assert.Equal(IrHasher.CanonicalHash(a), IrHasher.CanonicalHash(b));
    }

    [Fact]
    public void Canonicalize_RsidAndPt14_Stripped()
    {
        var bare = new XElement(W + "p",
            new XElement(W + "r", new XElement(W + "t", "hello")));

        var noisy = new XElement(W + "p",
            new XAttribute(W + "rsidR", "00AB12CD"),
            new XAttribute(Pt + "Unid", "deadbeef"),
            new XElement(W + "proofErr", new XAttribute(W + "type", "spellStart")),
            new XElement(W + "r", new XElement(W + "t", "hello")));

        Assert.Equal(IrHasher.CanonicalHash(bare), IrHasher.CanonicalHash(noisy));
    }

    [Fact]
    public void Canonicalize_DrawingAnchorAndEditIds_Stripped()
    {
        XElement Drawing(string? anchorId, string? editId, string docPrId, string cx = "100")
        {
            var inline = new XElement(Wp + "inline");
            if (anchorId is not null)
                inline.Add(new XAttribute(Wp14 + "anchorId", anchorId));
            if (editId is not null)
                inline.Add(new XAttribute(Wp14 + "editId", editId));
            inline.Add(
                new XElement(Wp + "docPr", new XAttribute("id", docPrId), new XAttribute("name", "Picture")),
                new XElement(Wp + "extent", new XAttribute("cx", cx), new XAttribute("cy", "200")));
            return new XElement(W + "drawing", inline);
        }

        var bare = Drawing(anchorId: null, editId: null, docPrId: "99");
        Assert.Equal(IrHasher.CanonicalHash(Drawing("5FC16A89", null, "1")), IrHasher.CanonicalHash(bare));
        Assert.Equal(IrHasher.CanonicalHash(Drawing(null, "1A20F538", "1")), IrHasher.CanonicalHash(bare));
        Assert.Equal(IrHasher.CanonicalHash(Drawing("5FC16A89", "1A20F538", "1")), IrHasher.CanonicalHash(bare));
        Assert.NotEqual(IrHasher.CanonicalHash(bare), IrHasher.CanonicalHash(Drawing(null, null, "99", cx: "101")));
    }

    [Fact]
    public void Canonicalize_ContentChange_Detected()
    {
        var a = new XElement(W + "p",
            new XElement(W + "r", new XElement(W + "t", "hello")));
        var b = new XElement(W + "p",
            new XElement(W + "r", new XElement(W + "t", "world")));

        Assert.NotEqual(IrHasher.CanonicalHash(a), IrHasher.CanonicalHash(b));
    }

    [Fact]
    public void Canonicalize_DoesNotMutateInput()
    {
        // Pt is the PowerTools (pt14) namespace the canonicalizer strips.
        var element = new XElement(W + "p",
            new XAttribute(W + "rsidR", "00AB12CD"),
            new XAttribute(Pt + "Unid", "deadbeef"),
            new XElement(W + "proofErr", new XAttribute(W + "type", "spellStart")),
            new XElement(W + "r", new XElement(W + "t", "hello")));

        IrHasher.Canonicalize(element);

        // The ORIGINAL element must still carry all three noise items.
        Assert.NotNull(element.Attribute(W + "rsidR"));
        Assert.NotNull(element.Attribute(Pt + "Unid"));
        Assert.NotNull(element.Element(W + "proofErr"));
    }

    // --- Content-hash stream builder --------------------------------------

    [Fact]
    public void ContentHash_TextVsSentinel_NoCollision()
    {
        // Tab SENTINEL must differ from the tab CHARACTER as text.
        var sentinel = new IrContentHashBuilder();
        sentinel.AppendText("a");
        sentinel.AppendSentinel(IrContentHashBuilder.SentinelTab);

        var asText = new IrContentHashBuilder();
        asText.AppendText("a\t");

        Assert.NotEqual(sentinel.Build(), asText.Build());
    }

    [Fact]
    public void ContentHash_StructureMarkers_Distinguish()
    {
        var withRow = new IrContentHashBuilder();
        withRow.AppendStructure(IrContentHashBuilder.StructureRow);
        withRow.AppendStructure(IrContentHashBuilder.StructureCell);
        withRow.AppendText("cell");

        var withoutRow = new IrContentHashBuilder();
        withoutRow.AppendStructure(IrContentHashBuilder.StructureCell);
        withoutRow.AppendText("cell");

        Assert.NotEqual(withRow.Build(), withoutRow.Build());
    }

    [Fact]
    public void ContentHash_AppendHash_Participates()
    {
        var h1 = IrHash.Compute("image-bytes-1");
        var h2 = IrHash.Compute("image-bytes-2");

        var a = new IrContentHashBuilder();
        a.AppendSentinel(IrContentHashBuilder.SentinelImage);
        a.AppendHash(h1);

        var b = new IrContentHashBuilder();
        b.AppendSentinel(IrContentHashBuilder.SentinelImage);
        b.AppendHash(h2);

        Assert.NotEqual(a.Build(), b.Build());
    }

    // --- Format fingerprints ----------------------------------------------

    [Fact]
    public void Fingerprint_NullFieldsOmitted()
    {
        var digest = IrHash.Compute("u");

        var boldOnly = new IrRunFormat { Bold = true, UnmodeledDigest = digest };
        var boldItalicNull = new IrRunFormat { Bold = true, Italic = null, UnmodeledDigest = digest };
        var italicOnly = new IrRunFormat { Italic = true, UnmodeledDigest = digest };

        Assert.Equal(
            IrHasher.FingerprintRunFormat(boldOnly),
            IrHasher.FingerprintRunFormat(boldItalicNull));
        Assert.NotEqual(
            IrHasher.FingerprintRunFormat(boldOnly),
            IrHasher.FingerprintRunFormat(italicOnly));
    }

    [Fact]
    public void Fingerprint_UnmodeledDigest_Participates()
    {
        var a = new IrRunFormat { Bold = true, UnmodeledDigest = IrHash.Compute("a") };
        var b = new IrRunFormat { Bold = true, UnmodeledDigest = IrHash.Compute("b") };

        Assert.NotEqual(IrHasher.FingerprintRunFormat(a), IrHasher.FingerprintRunFormat(b));
    }

    [Fact]
    public void Fingerprint_Deterministic_AcrossCalls()
    {
        var run = new IrRunFormat
        {
            Bold = true,
            Italic = false,
            Underline = new IrUnderline(IrUnderlineKind.Single, "FF0000"),
            SizeHalfPoints = 24,
            ColorHex = "auto",
            VertAlign = IrVertAlign.Superscript,
            UnmodeledDigest = IrHash.Compute("u"),
        };

        Assert.Equal(
            IrHasher.FingerprintRunFormat(run).ToHex(),
            IrHasher.FingerprintRunFormat(run).ToHex());

        var para = new IrParaFormat
        {
            Justification = IrJustification.Center,
            IndentLeftTwips = 720,
            LineSpacing = new IrLineSpacing(240, IrLineSpacingRule.Auto),
            UnmodeledDigest = IrHash.Compute("u"),
        };
        Assert.Equal(
            IrHasher.FingerprintParaFormat(para).ToHex(),
            IrHasher.FingerprintParaFormat(para).ToHex());
    }

    [Fact]
    public void Fingerprint_SectionFormat_Deterministic()
    {
        var sec = new IrSectionFormat
        {
            PageWidthTwips = 12240,
            PageHeightTwips = 15840,
            Landscape = false,
            SectionType = "nextPage",
            UnmodeledDigest = IrHash.Compute("u"),
        };

        Assert.Equal(
            IrHasher.FingerprintSectionFormat(sec).ToHex(),
            IrHasher.FingerprintSectionFormat(sec).ToHex());

        var sec2 = sec with { Landscape = true };
        Assert.NotEqual(
            IrHasher.FingerprintSectionFormat(sec),
            IrHasher.FingerprintSectionFormat(sec2));
    }

    [Fact]
    public void FingerprintBlock_CombinesParaAndRunFormats()
    {
        var para = new IrParaFormat { Justification = IrJustification.Left, UnmodeledDigest = IrHash.Compute("p") };
        var run1 = new IrRunFormat { Bold = true, UnmodeledDigest = IrHash.Compute("r1") };
        var run2 = new IrRunFormat { Italic = true, UnmodeledDigest = IrHash.Compute("r2") };

        var fpA = IrHasher.FingerprintBlock(para, new[] { run1, run2 });
        var fpB = IrHasher.FingerprintBlock(para, new[] { run1, run2 });
        Assert.Equal(fpA, fpB);

        // Run order matters.
        var fpReordered = IrHasher.FingerprintBlock(para, new[] { run2, run1 });
        Assert.NotEqual(fpA, fpReordered);

        // A bolded run flips the block fingerprint.
        var fpNoRuns = IrHasher.FingerprintBlock(para, new[] { run2 });
        Assert.NotEqual(fpA, fpNoRuns);
    }

    [Fact]
    public void Fingerprint_FieldFraming_Unambiguous()
    {
        var digest = IrHash.Compute("u");

        // StyleId value chosen to mimic the concatenation of "A" + "Bold=true" without framing.
        var packed = new IrRunFormat { StyleId = "A;Bold=true:1", UnmodeledDigest = digest };
        var split = new IrRunFormat { StyleId = "A", Bold = true, UnmodeledDigest = digest };

        Assert.NotEqual(
            IrHasher.FingerprintRunFormat(packed),
            IrHasher.FingerprintRunFormat(split));
    }

    // --- Every-field-flips-the-fingerprint guards -------------------------
    //
    // These catch the "added a property to the record but forgot to serialize it in
    // Fingerprint*" class of bugs: each single-field variant must differ from the
    // all-null baseline AND be pairwise-distinct from every other single-field variant.

    [Fact]
    public void Fingerprint_EveryRunFormatField_FlipsFingerprint()
    {
        var digest = IrHash.Compute("baseline");
        var baseline = new IrRunFormat { UnmodeledDigest = digest };

        var variants = new (string Name, IrRunFormat Value)[]
        {
            ("StyleId", baseline with { StyleId = "Heading1" }),
            ("Bold", baseline with { Bold = true }),
            ("Italic", baseline with { Italic = true }),
            ("Underline", baseline with { Underline = new IrUnderline(IrUnderlineKind.Single, "FF0000") }),
            ("Strike", baseline with { Strike = true }),
            ("DoubleStrike", baseline with { DoubleStrike = true }),
            ("VertAlign", baseline with { VertAlign = IrVertAlign.Superscript }),
            ("FontAscii", baseline with { FontAscii = "Calibri" }),
            ("SizeHalfPoints", baseline with { SizeHalfPoints = 24 }),
            ("ColorHex", baseline with { ColorHex = "FF0000" }),
            ("Highlight", baseline with { Highlight = "yellow" }),
            ("Shading", baseline with { Shading = new IrShading("<w:shd w:fill=\"FFFF00\"/>") }),
            ("Caps", baseline with { Caps = true }),
            ("SmallCaps", baseline with { SmallCaps = true }),
            ("Vanish", baseline with { Vanish = true }),
        };

        AssertPairwiseDistinctFromBaseline(
            IrHasher.FingerprintRunFormat(baseline),
            variants.Select(v => (v.Name, IrHasher.FingerprintRunFormat(v.Value))));
    }

    [Fact]
    public void Fingerprint_EveryParaFormatField_FlipsFingerprint()
    {
        var digest = IrHash.Compute("baseline");
        var baseline = new IrParaFormat { UnmodeledDigest = digest };

        var variants = new (string Name, IrParaFormat Value)[]
        {
            ("StyleId", baseline with { StyleId = "Normal" }),
            ("Justification", baseline with { Justification = IrJustification.Center }),
            ("IndentLeftTwips", baseline with { IndentLeftTwips = 720 }),
            ("IndentRightTwips", baseline with { IndentRightTwips = 720 }),
            ("IndentFirstLineTwips", baseline with { IndentFirstLineTwips = 360 }),
            ("SpacingBeforeTwips", baseline with { SpacingBeforeTwips = 120 }),
            ("SpacingAfterTwips", baseline with { SpacingAfterTwips = 120 }),
            ("LineSpacing", baseline with { LineSpacing = new IrLineSpacing(240, IrLineSpacingRule.Auto) }),
            ("OutlineLevel", baseline with { OutlineLevel = 1 }),
            ("KeepNext", baseline with { KeepNext = true }),
            ("KeepLines", baseline with { KeepLines = true }),
            ("PageBreakBefore", baseline with { PageBreakBefore = true }),
            ("Shading", baseline with { Shading = new IrShading("<w:shd w:fill=\"FFFF00\"/>") }),
        };

        AssertPairwiseDistinctFromBaseline(
            IrHasher.FingerprintParaFormat(baseline),
            variants.Select(v => (v.Name, IrHasher.FingerprintParaFormat(v.Value))));
    }

    [Fact]
    public void Fingerprint_EverySectionFormatField_FlipsFingerprint()
    {
        var digest = IrHash.Compute("baseline");
        var baseline = new IrSectionFormat { UnmodeledDigest = digest };

        var variants = new (string Name, IrSectionFormat Value)[]
        {
            ("PageWidthTwips", baseline with { PageWidthTwips = 12240 }),
            ("PageHeightTwips", baseline with { PageHeightTwips = 15840 }),
            ("Landscape", baseline with { Landscape = true }),
            ("MarginTopTwips", baseline with { MarginTopTwips = 1440 }),
            ("MarginBottomTwips", baseline with { MarginBottomTwips = 1440 }),
            ("MarginLeftTwips", baseline with { MarginLeftTwips = 1440 }),
            ("MarginRightTwips", baseline with { MarginRightTwips = 1440 }),
            ("SectionType", baseline with { SectionType = "nextPage" }),
        };

        AssertPairwiseDistinctFromBaseline(
            IrHasher.FingerprintSectionFormat(baseline),
            variants.Select(v => (v.Name, IrHasher.FingerprintSectionFormat(v.Value))));
    }

    private static void AssertPairwiseDistinctFromBaseline(
        IrHash baseline,
        IEnumerable<(string Name, IrHash Fingerprint)> variants)
    {
        var list = variants.ToList();
        var seen = new Dictionary<IrHash, string>();

        foreach (var (name, fp) in list)
        {
            Assert.True(fp != baseline, $"Field '{name}' did not flip the fingerprint vs baseline.");
            if (seen.TryGetValue(fp, out var other))
                Assert.Fail($"Field '{name}' collides with field '{other}': same fingerprint.");
            seen.Add(fp, name);
        }
    }
}
