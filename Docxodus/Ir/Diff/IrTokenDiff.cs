#nullable enable

using Docxodus.Ir;

namespace Docxodus.Ir.Diff;

/// <summary>
/// The kind of an intra-block token operation produced by <see cref="IrTokenDiffer"/> (M2.2 Task 1).
/// </summary>
internal enum IrTokenOpKind
{
    /// <summary>
    /// A run of tokens equal on both sides (equal <see cref="IrDiffToken.MatchKey"/>) AND with equal
    /// per-token <see cref="IrRunFormat"/> records. Left and right spans have the same length.
    /// </summary>
    Equal,

    /// <summary>
    /// A run of right-only tokens (present in <c>right</c>, absent in <c>left</c>). The left span is
    /// empty (<c>LeftStart == LeftEnd</c>) and anchors the insertion position in left coordinates.
    /// </summary>
    Insert,

    /// <summary>
    /// A run of left-only tokens (present in <c>left</c>, absent in <c>right</c>). The right span is
    /// empty (<c>RightStart == RightEnd</c>) and anchors the deletion position in right coordinates.
    /// </summary>
    Delete,

    /// <summary>
    /// A run of tokens equal in content (equal <see cref="IrDiffToken.MatchKey"/>) but differing in
    /// per-token <see cref="IrRunFormat"/>. Left and right spans have the same length. INVARIANT: every
    /// position inside the span has a pairwise-UNEQUAL format record (left[i].Format != right[i].Format
    /// by record equality, for every i in the span) — equal-format positions stay <see cref="Equal"/>,
    /// so a <see cref="FormatChanged"/> span is a maximal run of consecutive format-differing positions.
    /// </summary>
    FormatChanged,
}

/// <summary>
/// One intra-block token operation: a half-open token-index span on each side. The char spans are
/// derivable from the source token lists (<c>left[LeftStart].StartChar .. left[LeftEnd-1].EndChar</c>
/// and likewise on the right) so they are not stored redundantly here.
/// </summary>
/// <remarks>
/// Span conventions by kind:
/// <list type="bullet">
/// <item><see cref="IrTokenOpKind.Equal"/> / <see cref="IrTokenOpKind.FormatChanged"/>: both spans
/// non-empty and equal-length; <c>left[LeftStart+k].MatchKey == right[RightStart+k].MatchKey</c> for
/// every k.</item>
/// <item><see cref="IrTokenOpKind.Insert"/>: <c>LeftStart == LeftEnd</c> (empty left span at the
/// anchored position); right span non-empty.</item>
/// <item><see cref="IrTokenOpKind.Delete"/>: <c>RightStart == RightEnd</c> (empty right span at the
/// anchored position); left span non-empty.</item>
/// </list>
/// </remarks>
internal sealed record IrTokenOp(
    IrTokenOpKind Kind, int LeftStart, int LeftEnd, int RightStart, int RightEnd)
{
    /// <summary>Number of left tokens covered by this op.</summary>
    public int LeftLength => LeftEnd - LeftStart;

    /// <summary>Number of right tokens covered by this op.</summary>
    public int RightLength => RightEnd - RightStart;
}

/// <summary>
/// The intra-block token diff of one Modified paragraph pair: an ordered list of
/// <see cref="IrTokenOp"/> covering both sides exactly once.
/// </summary>
/// <remarks>
/// Totality: the left spans, concatenated in <see cref="Ops"/> order, tile <c>[0, leftCount)</c>
/// exactly once ascending; likewise the right spans tile <c>[0, rightCount)</c>. See
/// <c>IrTokenDiffAsserts</c> (test side) for the full invariant battery.
/// </remarks>
internal sealed record IrTokenDiff(IrNodeList<IrTokenOp> Ops);
