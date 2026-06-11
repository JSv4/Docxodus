#nullable enable

using System.Collections.Generic;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Shared invariant battery for <see cref="IrTokenDiffer"/> output, the token-diff counterpart to
/// <see cref="IrAlignmentAsserts"/>. Every <see cref="IrTokenDifferTests"/> case runs this against the
/// diff it produced, so the totality / coverage / per-kind invariants the plan pins are checked
/// uniformly rather than ad hoc per test.
/// </summary>
internal static class IrTokenDiffAsserts
{
    /// <summary>
    /// The token-diff invariants the plan pins:
    /// <list type="bullet">
    /// <item>Span well-formedness: <c>0 &lt;= Start &lt;= End</c> on both sides; Insert ⇒ empty left
    /// span; Delete ⇒ empty right span; Equal/FormatChanged ⇒ both spans non-empty and equal-length.</item>
    /// <item>Coverage: left spans, concatenated in op order, tile <c>[0, left.Count)</c> exactly once
    /// ascending (no gap, no overlap); right spans tile <c>[0, right.Count)</c> likewise.</item>
    /// <item>Equal/FormatChanged ⇒ pairwise-equal MatchKeys across the span.</item>
    /// <item>Equal ⇒ pairwise format-EQUAL under the policy; FormatChanged ⇒ pairwise format-UNEQUAL at
    /// EVERY position (a FormatChanged span is a maximal run of format-differing positions). "Format
    /// equal" follows <see cref="IrFormatComparison"/>: modeled-only (default) ignores unmodeled rPr
    /// noise, Full uses full record equality.</item>
    /// <item>No two adjacent ops share a coalescible kind boundary that should have merged: adjacent
    /// Equal-Equal, Insert-Insert, Delete-Delete are forbidden (they would not be maximal). Adjacent
    /// FormatChanged-FormatChanged and Equal-FormatChanged are allowed (the format pass produces
    /// maximal alternation, but two FormatChanged spans are only adjacent across an Equal, so we also
    /// forbid FormatChanged-FormatChanged).</item>
    /// </list>
    /// </summary>
    public static void AssertInvariants(
        IReadOnlyList<IrDiffToken> left, IReadOnlyList<IrDiffToken> right, IrTokenDiff diff,
        IrDiffSettings? settings = null)
    {
        var comparison = (settings ?? new IrDiffSettings()).FormatComparison;
        int leftCursor = 0, rightCursor = 0;
        IrTokenOp? prev = null;

        foreach (var op in diff.Ops)
        {
            // Well-formed spans.
            Assert.True(op.LeftStart >= 0 && op.LeftEnd >= op.LeftStart, "left span malformed");
            Assert.True(op.RightStart >= 0 && op.RightEnd >= op.RightStart, "right span malformed");

            // Coverage: each span must start exactly where the previous left off, on both sides.
            Assert.Equal(leftCursor, op.LeftStart);
            Assert.Equal(rightCursor, op.RightStart);

            switch (op.Kind)
            {
                case IrTokenOpKind.Insert:
                    Assert.Equal(op.LeftStart, op.LeftEnd);        // empty left span
                    Assert.True(op.RightLength > 0, "Insert must cover >=1 right token");
                    break;

                case IrTokenOpKind.Delete:
                    Assert.Equal(op.RightStart, op.RightEnd);      // empty right span
                    Assert.True(op.LeftLength > 0, "Delete must cover >=1 left token");
                    break;

                case IrTokenOpKind.Equal:
                case IrTokenOpKind.FormatChanged:
                    Assert.True(op.LeftLength > 0, "Equal/FormatChanged must cover >=1 token");
                    Assert.Equal(op.LeftLength, op.RightLength);   // equal-length
                    for (int k = 0; k < op.LeftLength; k++)
                    {
                        var l = left[op.LeftStart + k];
                        var r = right[op.RightStart + k];
                        Assert.Equal(l.MatchKey, r.MatchKey);      // pairwise-equal MatchKeys
                        bool fmtEqual = IrModeledFormat.RunFormatEqual(l.Format, r.Format, comparison);
                        if (op.Kind == IrTokenOpKind.Equal)
                            Assert.True(fmtEqual, "Equal position must have equal Format records");
                        else
                            Assert.False(fmtEqual, "FormatChanged position must have unequal Format records");
                    }
                    break;
            }

            // Maximality / no spurious adjacency.
            if (prev is not null)
            {
                bool sameMaximalKind =
                    (prev.Kind == op.Kind) &&
                    (op.Kind == IrTokenOpKind.Equal || op.Kind == IrTokenOpKind.Insert ||
                     op.Kind == IrTokenOpKind.Delete || op.Kind == IrTokenOpKind.FormatChanged);
                Assert.False(sameMaximalKind,
                    $"adjacent {op.Kind} ops must have been coalesced into one maximal span");
            }

            leftCursor = op.LeftEnd;
            rightCursor = op.RightEnd;
            prev = op;
        }

        // Totality: spans tiled the whole of both sides.
        Assert.Equal(left.Count, leftCursor);
        Assert.Equal(right.Count, rightCursor);
    }
}
