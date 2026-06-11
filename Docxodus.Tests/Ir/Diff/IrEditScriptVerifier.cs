#nullable enable

using System.Collections.Generic;
using System.Linq;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Apply-verifier for <see cref="IrEditScript"/> (M2.2 Task 2 exit invariant). Reconstructs the RIGHT
/// body's per-block token-text sequence from the LEFT IR + the script, then asserts it text-equals the
/// actual right body block-by-block. This is the "apply(script, left) reconstructs right at text level"
/// invariant the program-plan owes.
/// </summary>
/// <remarks>
/// <para><b>Why the verifier may consult the RIGHT document.</b> The invariant being proven is the
/// <em>structural consistency</em> of the script — that its anchors resolve, its ops are ordered so the
/// right-producing ops list the right blocks in right-document order, and its per-block token diffs map
/// the left content onto the right content faithfully — NOT that the script is a self-contained patch
/// carrying every inserted byte. So the verifier legitimately reads inserted content (InsertBlock) and
/// the right-side tokens of a ModifyBlock from the right IR: the script tells us WHERE and HOW MUCH to
/// take, the right IR supplies the literal tokens. A self-contained-bytes patch is an M2.3 renderer
/// concern, explicitly out of scope here.</para>
/// <para><b>Token text = MatchKey sequence.</b> We reconstruct and compare the normalized token
/// <see cref="IrDiffToken.MatchKey"/> sequence (the deterministic text view the diff keyed on), so a
/// case-insensitive / NBSP-folding settings choice is honored consistently on both sides. Non-paragraph
/// blocks (tables / opaque / section breaks) carry no token model in Task 2, so they are compared by
/// <see cref="IrBlock.ContentHash"/> instead.</para>
/// </remarks>
internal static class IrEditScriptVerifier
{
    private static readonly IrDiffSettings DefaultSettings = new();

    /// <summary>
    /// Verify the script transforms <paramref name="left"/>'s body into <paramref name="right"/>'s body
    /// at the text level. Asserts (1) every LeftAnchor resolves in left.AnchorIndex and every RightAnchor
    /// in right.AnchorIndex; (2) move source/destination pairing is well-formed; (3) the reconstructed
    /// right block sequence text-equals the actual right body.
    /// </summary>
    public static void Verify(IrDocument left, IrDocument right, IrEditScript script)
        => Verify(left, right, script, DefaultSettings);

    public static void Verify(IrDocument left, IrDocument right, IrEditScript script, IrDiffSettings settings)
    {
        AssertAnchorsResolve(left, right, script);
        AssertMovePairing(script);

        // Source blocks of moves, keyed by group, so a destination op can reproduce the moved-from text.
        var moveSourceBlock = new Dictionary<int, IrBlock>();
        foreach (var op in script.Operations)
        {
            if (op.IsMoveSource == true)
                moveSourceBlock[op.MoveGroupId!.Value] = ResolveLeft(left, op.LeftAnchor!);
        }

        // Each right-producing op contributes (the actual right block, the reconstructed paragraph
        // tokens OR null for a non-paragraph block, and the SOURCE block whose content the op claims
        // reproduces the right block — left/source for Equal/FormatOnly/Move, the right block itself for
        // Insert (inserted content legitimately IS the right block) and non-paragraph Modify (no sub-block
        // model in Task 2)). The SourceBlock drives a NON-VACUOUS ContentHash check for non-paragraph
        // EqualBlock/Move ops: comparing the left/source block's hash to the right block's hash proves the
        // op's content claim, instead of comparing the right block to itself.
        var reconstructed = new List<(IrBlock RightBlock, IReadOnlyList<string>? Tokens, IrBlock SourceBlock)>();

        foreach (var op in script.Operations)
        {
            switch (op.Kind)
            {
                case IrEditOpKind.EqualBlock:
                case IrEditOpKind.FormatOnlyBlock:
                {
                    // Unchanged text: the reconstructed right block is the left block's content.
                    var leftBlock = ResolveLeft(left, op.LeftAnchor!);
                    var rightBlock = ResolveRight(right, op.RightAnchor!);
                    reconstructed.Add((rightBlock, TokensOrNull(leftBlock, settings), leftBlock));
                    break;
                }

                case IrEditOpKind.ModifyBlock:
                {
                    var leftBlock = ResolveLeft(left, op.LeftAnchor!);
                    var rightBlock = ResolveRight(right, op.RightAnchor!);
                    // Non-paragraph Modify has no token model in Task 2; its content genuinely differs
                    // (it's Modified), so the source-of-truth for the non-paragraph check is the right block.
                    reconstructed.Add((rightBlock, ApplyModify(leftBlock, rightBlock, op.TokenDiff, settings), rightBlock));
                    break;
                }

                case IrEditOpKind.InsertBlock:
                {
                    // Inserted content comes from the right IR (legitimate; see remarks).
                    var rightBlock = ResolveRight(right, op.RightAnchor!);
                    reconstructed.Add((rightBlock, TokensOrNull(rightBlock, settings), rightBlock));
                    break;
                }

                case IrEditOpKind.MoveBlock:
                case IrEditOpKind.MoveModifyBlock:
                {
                    if (op.IsMoveSource == true)
                        break; // sources produce nothing on the right

                    var sourceBlock = moveSourceBlock[op.MoveGroupId!.Value];
                    var rightBlock = ResolveRight(right, op.RightAnchor!);
                    if (op.Kind == IrEditOpKind.MoveModifyBlock)
                        // MoveModify edits in flight; the right block is the content source-of-truth.
                        reconstructed.Add((rightBlock, ApplyModify(sourceBlock, rightBlock, op.TokenDiff, settings), rightBlock));
                    else
                        // Exact-content move: the destination reproduces the SOURCE text verbatim, so the
                        // source block drives both the paragraph token check and the non-paragraph hash check.
                        reconstructed.Add((rightBlock, TokensOrNull(sourceBlock, settings), sourceBlock));
                    break;
                }

                case IrEditOpKind.DeleteBlock:
                    break; // produces nothing on the right
            }
        }

        // The reconstructed right-producing ops must list the right blocks in right-document order.
        var actualRight = right.Body.Blocks;
        Assert.Equal(actualRight.Count, reconstructed.Count);
        for (int i = 0; i < actualRight.Count; i++)
        {
            var actual = actualRight[i];
            var (rightBlock, tokens, sourceBlock) = reconstructed[i];

            // The op named the i-th right block in order (reference identity).
            Assert.True(ReferenceEquals(actual, rightBlock),
                $"reconstructed right block #{i} ({rightBlock.Anchor}) is not the actual right block ({actual.Anchor}).");

            if (tokens is not null)
            {
                // Paragraph: reconstructed text must equal the actual right paragraph's text. We compare
                // the CONCATENATED MatchKey string, not the token-by-token sequence, because tokenization
                // BOUNDARIES are run-structure-dependent while the diff/aligner key on ContentHash (which
                // is boundary-independent). Concretely: a word abutting a non-separator across two runs
                // (e.g. "vil" + "»" split across runs on one side, "vil»" in one run on the other) yields
                // a DIFFERENT token COUNT on the two sides even though the text is identical and the
                // blocks are ContentHash-equal. Comparing the concatenation collapses that benign
                // run-boundary difference, so the verifier proves TEXT equality (the plan's "text level"
                // invariant) rather than over-asserting token-boundary identity the diff never claimed.
                var actualText = string.Concat(Tokens((IrParagraph)actual, settings));
                var reconstructedText = string.Concat(tokens);
                Assert.True(reconstructedText == actualText,
                    $"reconstructed paragraph #{i} ({actual.Anchor}) text mismatch:\n" +
                    $"  expected: [{actualText}]\n" +
                    $"  actual:   [{reconstructedText}]");
            }
            else
            {
                // Non-paragraph block: compare by ContentHash (no token model in Task 2). We compare the
                // SOURCE block's hash to the actual right block's hash — for an EqualBlock/MoveBlock the
                // source is the left/moved-from block, so this NON-VACUOUSLY proves the op reproduced the
                // right block's content (a mislabeled EqualBlock over two differing tables would fail here).
                // For Insert / non-paragraph Modify the source IS the right block (content legitimately
                // sourced from the right IR), so the check is identity there by design.
                Assert.Equal(actual.ContentHash, sourceBlock.ContentHash);
            }
        }
    }

    // ------------------------------------------------------------------ helpers

    private static void AssertAnchorsResolve(IrDocument left, IrDocument right, IrEditScript script)
    {
        foreach (var op in script.Operations)
        {
            if (op.LeftAnchor is { } la)
                Assert.True(left.AnchorIndex.ContainsKey(la), $"LeftAnchor '{la}' does not resolve in left.AnchorIndex.");
            if (op.RightAnchor is { } ra)
                Assert.True(right.AnchorIndex.ContainsKey(ra), $"RightAnchor '{ra}' does not resolve in right.AnchorIndex.");
        }
    }

    /// <summary>
    /// Each move group must have EXACTLY one source op (IsMoveSource=true, LeftAnchor set, no RightAnchor)
    /// and exactly one destination op (IsMoveSource=false, RightAnchor set, no LeftAnchor), sharing the
    /// op kind. Group ids are 1..N contiguous and assigned in destination order.
    /// </summary>
    private static void AssertMovePairing(IrEditScript script)
    {
        var sources = new Dictionary<int, IrEditOp>();
        var destinations = new Dictionary<int, IrEditOp>();
        var destinationOrder = new List<int>();

        foreach (var op in script.Operations)
        {
            if (op.Kind is not (IrEditOpKind.MoveBlock or IrEditOpKind.MoveModifyBlock))
            {
                Assert.Null(op.MoveGroupId);
                Assert.Null(op.IsMoveSource);
                continue;
            }

            Assert.NotNull(op.MoveGroupId);
            Assert.NotNull(op.IsMoveSource);
            int group = op.MoveGroupId!.Value;

            if (op.IsMoveSource == true)
            {
                Assert.NotNull(op.LeftAnchor);
                Assert.Null(op.RightAnchor);
                Assert.Null(op.TokenDiff); // the diff lives on the destination, not the source
                Assert.False(sources.ContainsKey(group), $"duplicate move source for group {group}.");
                sources[group] = op;
            }
            else
            {
                Assert.Null(op.LeftAnchor);
                Assert.NotNull(op.RightAnchor);
                // A MoveModify DESTINATION must carry its in-move token diff; a plain Move destination must not.
                if (op.Kind == IrEditOpKind.MoveModifyBlock)
                    Assert.NotNull(op.TokenDiff);
                else
                    Assert.Null(op.TokenDiff);
                Assert.False(destinations.ContainsKey(group), $"duplicate move destination for group {group}.");
                destinations[group] = op;
                destinationOrder.Add(group);
            }
        }

        Assert.Equal(sources.Count, destinations.Count);
        foreach (var (group, dest) in destinations)
        {
            Assert.True(sources.TryGetValue(group, out var src), $"move group {group} has a destination but no source.");
            Assert.Equal(src!.Kind, dest.Kind); // source/destination share the op kind
        }

        // Group ids: 1..N contiguous, assigned in destination order.
        Assert.Equal(Enumerable.Range(1, destinationOrder.Count).ToList(), destinationOrder);
    }

    private static IReadOnlyList<string>? TokensOrNull(IrBlock block, IrDiffSettings settings) =>
        block is IrParagraph p ? Tokens(p, settings) : null;

    private static IReadOnlyList<string> Tokens(IrParagraph p, IrDiffSettings settings) =>
        IrDiffTokenizer.Tokenize(p, settings).Select(t => t.MatchKey).ToList();

    /// <summary>
    /// Reconstruct the right paragraph's token sequence by applying <paramref name="tokenDiff"/> to the
    /// left tokens: Equal/FormatChanged copy left tokens; Insert takes right tokens (from the right IR);
    /// Delete drops left tokens. For a non-paragraph Modified pair (null diff) returns null so the caller
    /// compares by ContentHash.
    /// </summary>
    private static IReadOnlyList<string>? ApplyModify(
        IrBlock leftBlock, IrBlock rightBlock, IrTokenDiff? tokenDiff, IrDiffSettings settings)
    {
        if (leftBlock is not IrParagraph lp || rightBlock is not IrParagraph rp)
            return null; // non-paragraph Modified: compared by ContentHash
        Assert.NotNull(tokenDiff); // a paragraph Modified pair MUST carry a token diff

        var leftTokens = IrDiffTokenizer.Tokenize(lp, settings);
        var rightTokens = IrDiffTokenizer.Tokenize(rp, settings);

        // Enforce the full TokenDiff totality/coverage/per-kind battery here too, so the apply-verifier
        // does NOT pass on a structurally-broken-but-text-equal diff (non-tiling spans, wrong-length
        // Equal/FormatChanged runs, corrupt right spans). This makes the concatenated-text check below a
        // SUFFICIENT-given-tiling proof rather than the sole assertion.
        IrTokenDiffAsserts.AssertInvariants(leftTokens, rightTokens, tokenDiff!);

        var result = new List<string>();
        foreach (var op in tokenDiff!.Ops)
        {
            switch (op.Kind)
            {
                case IrTokenOpKind.Equal:
                case IrTokenOpKind.FormatChanged:
                    for (int k = op.LeftStart; k < op.LeftEnd; k++)
                        result.Add(leftTokens[k].MatchKey);
                    break;
                case IrTokenOpKind.Insert:
                    for (int k = op.RightStart; k < op.RightEnd; k++)
                        result.Add(rightTokens[k].MatchKey);
                    break;
                case IrTokenOpKind.Delete:
                    break; // left tokens dropped
            }
        }
        return result;
    }

    private static IrBlock ResolveLeft(IrDocument left, string anchor)
    {
        Assert.True(left.AnchorIndex.TryGetValue(anchor, out var block), $"LeftAnchor '{anchor}' missing.");
        return block!;
    }

    private static IrBlock ResolveRight(IrDocument right, string anchor)
    {
        Assert.True(right.AnchorIndex.TryGetValue(anchor, out var block), $"RightAnchor '{anchor}' missing.");
        return block!;
    }
}
