#nullable enable

using System.Linq;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// M2.2 Task 4 sub-task B tests for the <see cref="IrFormatComparison"/> policy. Two content-equal
/// paragraphs differing ONLY in a residual rPr child (the WC-BodyBookmarks noise pattern: w:lang /
/// w:bCs / w:iCs / w:szCs / w:rFonts cs faces) must classify as Unchanged under the default ModeledOnly
/// policy (the residual digest flips the stored FormatFingerprint but no MODELED field changed) and as
/// FormatOnly under Full. A visible `w:shd` change is deliberately modeled under BOTH policies.
/// </summary>
public class IrFormatComparisonTests
{
    private static readonly IrReaderOptions NoSources = new() { RetainSources = false };

    private static IrDocument FromXml(string bodyInnerXml) =>
        IrReader.Read(IrTestDocuments.FromBodyXml(bodyInnerXml), NoSources);

    // A paragraph with one run carrying the given rPr inner XML around the text "hello world".
    private static string Para(string rPrInner) =>
        $"<w:p><w:r><w:rPr>{rPrInner}</w:rPr><w:t>hello world</w:t></w:r></w:p>";

    private static IrBlockAlignment Align(IrDocument l, IrDocument r, IrFormatComparison cmp) =>
        IrBlockAligner.Align(l, r, new IrDiffSettings { FormatComparison = cmp });

    [Theory]
    [InlineData("<w:lang w:val=\"nb-NO\"/>")]   // locale — the dominant BodyBookmarks noise (4597 occ.)
    [InlineData("<w:bCs/>")]                      // complex-script bold toggle (550 occ.)
    [InlineData("<w:iCs/>")]                      // complex-script italic toggle (1328 occ.)
    [InlineData("<w:szCs w:val=\"24\"/>")]       // complex-script size (3 occ.)
    public void Residual_rpr_only_diff_is_unchanged_under_modeled_only_but_formatonly_under_full(string noiseRpr)
    {
        var left = FromXml(Para(string.Empty));
        var right = FromXml(Para(noiseRpr));

        // Sanity: content equal, but the stored FormatFingerprint flipped (the unmodeled digest changed).
        var lb = left.Body.Blocks[0];
        var rb = right.Body.Blocks[0];
        Assert.Equal(lb.ContentHash, rb.ContentHash);
        Assert.NotEqual(lb.FormatFingerprint, rb.FormatFingerprint);

        // ModeledOnly (default): no modeled field changed ⇒ Unchanged.
        var modeled = Align(left, right, IrFormatComparison.ModeledOnly);
        Assert.Equal(IrAlignmentKind.Unchanged, modeled.Entries.Single().Kind);
        IrAlignmentAsserts.AssertInvariants(left, right, modeled,
            new IrDiffSettings { FormatComparison = IrFormatComparison.ModeledOnly });

        // Full: the stored fingerprint difference is honored ⇒ FormatOnly.
        var full = Align(left, right, IrFormatComparison.Full);
        Assert.Equal(IrAlignmentKind.FormatOnly, full.Entries.Single().Kind);
        IrAlignmentAsserts.AssertInvariants(left, right, full,
            new IrDiffSettings { FormatComparison = IrFormatComparison.Full });
    }

    [Fact]
    public void Shading_only_rpr_diff_is_formatonly_under_both_policies()
    {
        var left = FromXml(Para(string.Empty));
        var right = FromXml(Para("<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"FFFF00\"/>"));

        foreach (var comparison in new[] { IrFormatComparison.ModeledOnly, IrFormatComparison.Full })
        {
            var alignment = Align(left, right, comparison);
            Assert.Equal(IrAlignmentKind.FormatOnly, alignment.Entries.Single().Kind);
            IrAlignmentAsserts.AssertInvariants(left, right, alignment,
                new IrDiffSettings { FormatComparison = comparison });
        }
    }

    [Fact]
    public void Modeled_format_change_is_formatonly_under_both_policies()
    {
        var left = FromXml(Para(string.Empty));
        var right = FromXml(Para("<w:b/>")); // genuine modeled bold change

        foreach (var cmp in new[] { IrFormatComparison.ModeledOnly, IrFormatComparison.Full })
        {
            var a = Align(left, right, cmp);
            Assert.Equal(IrAlignmentKind.FormatOnly, a.Entries.Single().Kind);
        }
    }

    [Fact]
    public void Token_diff_ignores_unmodeled_noise_under_modeled_only()
    {
        // Same text, the second word's run carries a lang-only rPr difference: under ModeledOnly the
        // token diff is all-Equal (no FormatChanged); under Full it raises a FormatChanged span.
        string LeftPara() =>
            "<w:p><w:r><w:t xml:space=\"preserve\">hello </w:t></w:r>" +
            "<w:r><w:t>world</w:t></w:r></w:p>";
        string RightPara() =>
            "<w:p><w:r><w:t xml:space=\"preserve\">hello </w:t></w:r>" +
            "<w:r><w:rPr><w:lang w:val=\"nb-NO\"/></w:rPr><w:t>world</w:t></w:r></w:p>";

        var lp = (IrParagraph)FromXml(LeftPara()).Body.Blocks[0];
        var rp = (IrParagraph)FromXml(RightPara()).Body.Blocks[0];

        var modeled = new IrDiffSettings { FormatComparison = IrFormatComparison.ModeledOnly };
        var lt = IrDiffTokenizer.Tokenize(lp, modeled);
        var rt = IrDiffTokenizer.Tokenize(rp, modeled);
        var diffModeled = IrTokenDiffer.Diff(lt, rt, modeled);
        Assert.DoesNotContain(diffModeled.Ops, o => o.Kind == IrTokenOpKind.FormatChanged);

        var full = new IrDiffSettings { FormatComparison = IrFormatComparison.Full };
        var diffFull = IrTokenDiffer.Diff(
            IrDiffTokenizer.Tokenize(lp, full), IrDiffTokenizer.Tokenize(rp, full), full);
        Assert.Contains(diffFull.Ops, o => o.Kind == IrTokenOpKind.FormatChanged);
    }
}
