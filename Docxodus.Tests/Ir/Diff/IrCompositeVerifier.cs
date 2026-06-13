#nullable enable

using System.Collections.Generic;
using System.Linq;
using Docxodus;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Apply-verifier for <see cref="IrCompositeScript"/> (T5.1) — the composite analogue of
/// <see cref="IrEditScriptVerifier"/>. It INDEPENDENTLY reconstructs the policy-resolved accepted body
/// text purely from the composite script's op semantics (applying the merged/authored ops to the base),
/// then asserts that text equals what the RENDERER produced —
/// <c>PlainText(AcceptRevisions(DocxDiff.Consolidate(base, reviewers, settings)))</c>. If the two
/// disagree the script ops do not faithfully describe the rendered output: a real bug, not a verifier
/// quirk. The verifier never mirrors the renderer; it derives the body from the ops alone and the
/// renderer-side text is the oracle it checks against.
/// </summary>
/// <remarks>
/// <para><b>Conflicts are pre-resolved into the op stream.</b> The merger has already applied the policy
/// (BaseWins → Equal at the conflicted span; FirstReviewerWins → the first reviewer's op; StackAll →
/// stacked ops), so applying the ops directly yields the policy-resolved body. The verifier therefore
/// needs no conflict-policy branching of its own.</para>
/// <para><b>Where content is sourced.</b> As in <see cref="IrEditScriptVerifier"/>, the script says
/// WHERE and HOW MUCH; the literal tokens come from the IRs. EqualBlock text comes from the base block;
/// an InsertBlock and a ModifyBlock's inserted tokens come from the CONTRIBUTING reviewer's IR (indexed
/// by <see cref="IrCompositeOp.SourceReviewer"/> / per-span
/// <see cref="IrAuthoredTokenOp.SourceReviewer"/> via <see cref="IrCompositeOp.SourceRightAnchors"/>).</para>
/// </remarks>
internal static class IrCompositeVerifier
{
    internal static readonly IrReaderOptions ReadOpts =
        new() { RetainSources = false, RevisionView = RevisionView.Accept };

    public static void Verify(
        WmlDocument baseDoc,
        IReadOnlyList<(string Author, WmlDocument Doc)> reviewers,
        ConflictResolution policy,
        string expectedAcceptedBodyText)
    {
        var diff = new DocxDiffSettings().ToIrDiffSettings();
        var baseIr = IrReader.Read(baseDoc, ReadOpts);
        var revIr = reviewers.Select(r => (r.Author, IrReader.Read(r.Doc, ReadOpts))).ToList();
        var script = IrCompositeMerger.Merge(baseIr, revIr, policy, diff);

        // One reconstructed text fragment per right-producing op, in op order.
        var blocks = new List<string>();
        foreach (var cop in script.Operations)
        {
            var op = cop.Op;
            switch (op.Kind)
            {
                case IrEditOpKind.EqualBlock:
                case IrEditOpKind.FormatOnlyBlock:
                    // Unchanged base block: its accepted token text (merger sets both anchors = base anchor).
                    blocks.Add(BaseBlockText(baseIr, op.LeftAnchor!, diff));
                    break;

                case IrEditOpKind.DeleteBlock:
                    break; // contributes nothing to the body

                case IrEditOpKind.InsertBlock:
                    // Inserted whole block: sourced from the contributing reviewer's right IR.
                    blocks.Add(ReviewerBlockText(revIr, cop.SourceReviewer, op.RightAnchor!, diff));
                    break;

                case IrEditOpKind.ModifyBlock:
                    blocks.Add(ReconstructModify(cop, baseIr, revIr, diff));
                    break;

                case IrEditOpKind.MoveBlock:
                case IrEditOpKind.MoveModifyBlock:
                    // Single-source verbatim move (the merger only emits these as a lone reviewer's op):
                    // the destination produces the reviewer's right block text; the source produces nothing.
                    if (op.IsMoveSource != true)
                        blocks.Add(ReviewerBlockText(revIr, cop.SourceReviewer, op.RightAnchor!, diff));
                    break;

                case IrEditOpKind.SplitBlock:
                    // One reviewer split a base paragraph into N right members → N text fragments in order.
                    foreach (var a in op.SplitMergeAnchors!)
                        blocks.Add(ReviewerBlockText(revIr, cop.SourceReviewer, a, diff));
                    break;

                case IrEditOpKind.MergeBlock:
                    // N base members collapsed into one reviewer right paragraph.
                    blocks.Add(ReviewerBlockText(revIr, cop.SourceReviewer, op.RightAnchor!, diff));
                    break;
            }
        }

        var reconstructed = RevisionEquivalence.Normalize(string.Join("\n", blocks));
        var expected = RevisionEquivalence.Normalize(expectedAcceptedBodyText);
        Assert.True(reconstructed == expected,
            "composite apply-reconstruction does not match the rendered accepted body:\n" +
            $"  expected (renderer): [{expected}]\n" +
            $"  actual   (script):   [{reconstructed}]");
    }

    /// <summary>
    /// Reconstruct a ModifyBlock's right text from the op alone. A single-source modify applies its own
    /// <see cref="IrEditOp.TokenDiff"/> to the base block tokens (reusing the edit-script verifier's
    /// apply). A COMPOSED multi-reviewer modify walks <see cref="IrCompositeOp.AuthoredTokens"/>:
    /// Equal/FormatChanged → base token text; Insert → that span's contributing reviewer's right token
    /// text (resolved via <see cref="IrCompositeOp.SourceRightAnchors"/>); Delete → nothing.
    /// </summary>
    private static string ReconstructModify(
        IrCompositeOp cop, IrDocument baseIr,
        IReadOnlyList<(string Author, IrDocument Ir)> revIr, IrDiffSettings settings)
    {
        var op = cop.Op;
        var baseBlock = Resolve(baseIr, op.LeftAnchor!);

        // Non-paragraph modify (e.g. a table): no token model — fall back to the right block's text.
        if (baseBlock is not IrParagraph)
            return ReviewerBlockText(revIr, cop.SourceReviewer, op.RightAnchor!, settings);

        if (cop.AuthoredTokens is not { } authored)
        {
            // Single-source modify: apply this reviewer's token diff to the base paragraph.
            var rightBlock = Resolve(revIr[cop.SourceReviewer].Ir, op.RightAnchor!);
            var tokens = IrEditScriptVerifier.ApplyModify(baseBlock, rightBlock, op.TokenDiff, settings);
            Assert.NotNull(tokens); // a paragraph modify must carry a token diff
            return string.Concat(tokens!);
        }

        // Composed multi-reviewer modify: walk the authored spans.
        var baseTokens = IrEditScriptVerifier.MaskedTokenize((IrParagraph)baseBlock, settings);
        // Per contributing reviewer, the right tokens its Insert spans index into.
        var rightTokensByReviewer = new Dictionary<int, IReadOnlyList<IrDiffToken>>();
        if (cop.SourceRightAnchors is { } sras)
            foreach (var sra in sras)
                rightTokensByReviewer[sra.Reviewer] =
                    IrEditScriptVerifier.MaskedTokenize((IrParagraph)Resolve(revIr[sra.Reviewer].Ir, sra.Anchor), settings);

        var result = new System.Text.StringBuilder();
        foreach (var at in authored)
        {
            var t = at.Op;
            switch (t.Kind)
            {
                case IrTokenOpKind.Equal:
                case IrTokenOpKind.FormatChanged:
                    for (int k = t.LeftStart; k < t.LeftEnd; k++)
                        result.Append(baseTokens[k].MatchKey);
                    break;
                case IrTokenOpKind.Insert:
                    Assert.True(rightTokensByReviewer.TryGetValue(at.SourceReviewer, out var rt),
                        $"composed Insert authored to reviewer {at.SourceReviewer} has no SourceRightAnchor.");
                    for (int k = t.RightStart; k < t.RightEnd; k++)
                        result.Append(rt![k].MatchKey);
                    break;
                case IrTokenOpKind.Delete:
                    break; // base token dropped
            }
        }
        return result.ToString();
    }

    private static string BaseBlockText(IrDocument baseIr, string anchor, IrDiffSettings settings) =>
        BlockText(Resolve(baseIr, anchor), settings);

    private static string ReviewerBlockText(
        IReadOnlyList<(string Author, IrDocument Ir)> revIr, int reviewer, string anchor, IrDiffSettings settings) =>
        BlockText(Resolve(revIr[reviewer].Ir, anchor), settings);

    private static string BlockText(IrBlock block, IrDiffSettings settings) =>
        block is IrParagraph p ? string.Concat(IrEditScriptVerifier.Tokens(p, settings)) : block.ContentHash.ToHex();

    private static IrBlock Resolve(IrDocument ir, string anchor)
    {
        Assert.True(ir.AnchorIndex.TryGetValue(anchor, out var block), $"anchor '{anchor}' does not resolve.");
        return block!;
    }
}
