#nullable enable

using System.Linq;
using Docxodus.Ir;
using Xunit;

namespace Docxodus.Tests.Ir;

/// <summary>
/// M1.2 normalization-rule tests (spec §5.2). One test per rule landed in this milestone task:
/// N3 (bookmarks), N4 (lastRenderedPageBreak), N7 (special hyphens), N8 (w:sym), and the
/// strip-half of N15 (comment plumbing). The invariant under test throughout: noise-only
/// differences must produce identical <c>ContentHash</c> and <c>FormatFingerprint</c>, while
/// genuine glyph content (a <c>w:sym</c>) is allowed to move both.
/// </summary>
public class IrNormalizationTests
{
    private static IrParagraph Para(string bodyXml) =>
        IrReader.Read(IrTestDocuments.FromBodyXml(bodyXml))
            .Body.Blocks.OfType<IrParagraph>().Single();

    private static string Text(IrParagraph p) =>
        string.Concat(p.Inlines.OfType<IrTextRun>().Select(r => r.Text));

    [Fact]
    public void Read_Bookmarks_DoNotAffectHashes()
    {
        var clean = Para("<w:p><w:r><w:t>hello</w:t></w:r></w:p>");
        var bookmarked = Para(
            "<w:p><w:bookmarkStart w:id=\"0\" w:name=\"_GoBack\"/>" +
            "<w:r><w:t>hello</w:t></w:r>" +
            "<w:bookmarkEnd w:id=\"0\"/></w:p>");

        Assert.Equal(clean.ContentHash.ToHex(), bookmarked.ContentHash.ToHex());
        Assert.Equal(clean.FormatFingerprint.ToHex(), bookmarked.FormatFingerprint.ToHex());
        Assert.DoesNotContain(bookmarked.Inlines, i => i is IrOpaqueInline);
    }

    [Fact]
    public void Read_BookmarksInsideRun_DoNotAffectHashes()
    {
        // Bookmarks can legally sit inside a run, not just at paragraph level (N3 covers both).
        var clean = Para("<w:p><w:r><w:t>hello</w:t></w:r></w:p>");
        var bookmarked = Para(
            "<w:p><w:r><w:bookmarkStart w:id=\"0\" w:name=\"b\"/>" +
            "<w:t>hello</w:t><w:bookmarkEnd w:id=\"0\"/></w:r></w:p>");

        Assert.Equal(clean.ContentHash.ToHex(), bookmarked.ContentHash.ToHex());
        Assert.Equal(clean.FormatFingerprint.ToHex(), bookmarked.FormatFingerprint.ToHex());
        Assert.DoesNotContain(bookmarked.Inlines, i => i is IrOpaqueInline);
    }

    [Fact]
    public void Read_LastRenderedPageBreak_Dropped()
    {
        var clean = Para("<w:p><w:r><w:t>a</w:t><w:t>b</w:t></w:r></w:p>");
        var withBreak = Para(
            "<w:p><w:r><w:t>a</w:t><w:lastRenderedPageBreak/><w:t>b</w:t></w:r></w:p>");

        Assert.Equal(clean.ContentHash.ToHex(), withBreak.ContentHash.ToHex());
        Assert.Equal(clean.FormatFingerprint.ToHex(), withBreak.FormatFingerprint.ToHex());
        Assert.DoesNotContain(withBreak.Inlines, i => i is IrOpaqueInline);
    }

    [Fact]
    public void Read_SpecialHyphens_BecomeText()
    {
        // a + noBreakHyphen + b in a single run → one coalesced run "a‑b".
        var p = Para("<w:p><w:r><w:t>a</w:t><w:noBreakHyphen/><w:t>b</w:t></w:r></w:p>");

        var run = Assert.Single(p.Inlines.OfType<IrTextRun>());
        Assert.Equal("a‑b", run.Text);
        Assert.DoesNotContain(p.Inlines, i => i is IrOpaqueInline);

        // soft hyphen maps to U+00AD with the same coalescing behavior.
        var soft = Para("<w:p><w:r><w:t>x</w:t><w:softHyphen/><w:t>y</w:t></w:r></w:p>");
        var softRun = Assert.Single(soft.Inlines.OfType<IrTextRun>());
        Assert.Equal("x­y", softRun.Text);
    }

    [Fact]
    public void Read_Sym_BecomesTextAndFlipsFingerprint()
    {
        var plain = Para("<w:p><w:r><w:t>plain</w:t></w:r></w:p>");
        var sym = Para(
            "<w:p><w:r><w:sym w:font=\"Wingdings\" w:char=\"F0E0\"/></w:r></w:p>");

        // The symbol becomes the literal code point U+F0E0 in text.
        Assert.Equal("", Text(sym));

        // It is genuine content + glyph formatting: both hashes move relative to plain text.
        Assert.NotEqual(plain.ContentHash.ToHex(), sym.ContentHash.ToHex());
        Assert.NotEqual(plain.FormatFingerprint.ToHex(), sym.FormatFingerprint.ToHex());
        Assert.DoesNotContain(sym.Inlines, i => i is IrOpaqueInline);

        // The font is part of the fingerprint: a different font flips it; the same char keeps
        // ContentHash stable (the glyph is the same code point).
        var symOtherFont = Para(
            "<w:p><w:r><w:sym w:font=\"Symbol\" w:char=\"F0E0\"/></w:r></w:p>");
        Assert.Equal(sym.ContentHash.ToHex(), symOtherFont.ContentHash.ToHex());
        Assert.NotEqual(sym.FormatFingerprint.ToHex(), symOtherFont.FormatFingerprint.ToHex());
    }

    [Fact]
    public void Read_Sym_Unparseable_StaysOpaque()
    {
        // No @w:char → cannot map to a code point → opaque fallback.
        var p = Para("<w:p><w:r><w:sym w:font=\"Wingdings\"/></w:r></w:p>");

        Assert.Empty(p.Inlines.OfType<IrTextRun>());
        var opaque = Assert.Single(p.Inlines.OfType<IrOpaqueInline>());
        Assert.Equal("sym", opaque.ElementName.LocalName);
    }

    [Fact]
    public void Read_CommentPlumbing_DoesNotAffectHashes()
    {
        var clean = Para("<w:p><w:r><w:t>commented</w:t></w:r></w:p>");
        var commented = Para(
            "<w:p><w:commentRangeStart w:id=\"0\"/>" +
            "<w:r><w:t>commented</w:t></w:r>" +
            "<w:commentRangeEnd w:id=\"0\"/>" +
            "<w:r><w:commentReference w:id=\"0\"/></w:r></w:p>");

        Assert.Equal(clean.ContentHash.ToHex(), commented.ContentHash.ToHex());
        Assert.Equal(clean.FormatFingerprint.ToHex(), commented.FormatFingerprint.ToHex());
        Assert.DoesNotContain(commented.Inlines, i => i is IrOpaqueInline);
    }
}
