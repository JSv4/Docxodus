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
                    // A Modified TABLE pair carries a nested table diff (M2.2 Task 4): validate its
                    // row/cell anchors + reconstruct the right table row-by-row/cell-by-cell from it.
                    if (leftBlock is IrTable lt && rightBlock is IrTable rt && op.TableDiff is { } tableDiff)
                        VerifyTableDiff(lt, rt, tableDiff, settings);
                    // A Modified PARAGRAPH pair carrying textbox diffs (M2.4 Task 1): validate each
                    // textbox's inner block ops reconstruct the right textbox's blocks.
                    if (leftBlock is IrParagraph lpar && rightBlock is IrParagraph rpar && op.TextboxDiffs is { } tbxDiffs)
                        VerifyTextboxDiffs(lpar, rpar, tbxDiffs, settings);
                    // Block-level: a non-paragraph Modify's content genuinely differs (it's Modified), so
                    // the block-sequence check sources truth from the right block (table reconstruction is
                    // proven in depth by VerifyTableDiff above).
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

        // Note scopes (M2.4 Task 1): every footnote/endnote diff reconstructs its matched note's right
        // blocks from the left note's blocks (or all-insert/all-delete for an unmatched note).
        VerifyNoteOps(left, right, script, settings);
    }

    // ------------------------------------------------------------------ note scopes (M2.4 Task 1)

    /// <summary>
    /// Verify the script's per-note diffs. For each <see cref="IrNoteDiff"/>: resolve its note id in the
    /// correct (footnote/endnote) store on each side; reconstruct the note's right block text sequence from
    /// the left note's blocks + the note's ops (reusing the cell-block reconstruction machinery) and assert
    /// it equals the right note's block text. An unmatched note has empty left or right blocks.
    /// </summary>
    private static void VerifyNoteOps(IrDocument left, IrDocument right, IrEditScript script, IrDiffSettings settings)
    {
        if (script.NoteOps is not { } noteOps)
            return;

        foreach (var noteDiff in noteOps)
        {
            var leftStore = noteDiff.Kind == IrNoteKind.Footnote ? left.Footnotes : left.Endnotes;
            var rightStore = noteDiff.Kind == IrNoteKind.Footnote ? right.Footnotes : right.Endnotes;

            var leftBlocks = leftStore.Notes.TryGetValue(noteDiff.NoteId, out var ls)
                ? ls.Blocks : IrNodeList.Empty<IrBlock>();
            var rightBlocks = rightStore.Notes.TryGetValue(noteDiff.NoteId, out var rs)
                ? rs.Blocks : IrNodeList.Empty<IrBlock>();

            // At least one side must hold the note (the diff would not exist otherwise).
            Assert.True(leftBlocks.Count > 0 || rightBlocks.Count > 0,
                $"note {noteDiff.Kind}:{noteDiff.NoteId} resolves to no blocks on either side.");

            var reconstructed = ReconstructBlocks(leftBlocks, rightBlocks, noteDiff.Ops, settings);
            var expected = BlockTexts(rightBlocks, settings);
            Assert.Equal(expected, reconstructed);
        }
    }

    // ------------------------------------------------------------------ textbox interiors (M2.4 Task 1)

    /// <summary>
    /// Verify a Modified paragraph's textbox diffs: collect each paragraph's textboxes in document order,
    /// pair them positionally, and assert each diff's inner block ops reconstruct the paired right textbox's
    /// block text (or the lone surplus textbox's all-insert/all-delete reconstruction).
    /// </summary>
    private static void VerifyTextboxDiffs(
        IrParagraph left, IrParagraph right, IrNodeList<IrTextboxDiff> diffs, IrDiffSettings settings)
    {
        var leftBoxes = CollectTextboxes(left.Inlines);
        var rightBoxes = CollectTextboxes(right.Inlines);

        // One diff per textbox slot: paired boxes first, then the surplus on whichever side is longer.
        Assert.Equal(Math.Max(leftBoxes.Count, rightBoxes.Count), diffs.Count);

        for (int i = 0; i < diffs.Count; i++)
        {
            var lb = i < leftBoxes.Count ? leftBoxes[i].Blocks : IrNodeList.Empty<IrBlock>();
            var rb = i < rightBoxes.Count ? rightBoxes[i].Blocks : IrNodeList.Empty<IrBlock>();
            var reconstructed = ReconstructBlocks(lb, rb, diffs[i].Ops, settings);
            var expected = BlockTexts(rb, settings);
            Assert.Equal(expected, reconstructed);
        }
    }

    private static List<IrTextbox> CollectTextboxes(IReadOnlyList<IrInline> inlines)
    {
        var boxes = new List<IrTextbox>();
        WalkForTextboxes(inlines, boxes);
        return boxes;
    }

    private static void WalkForTextboxes(IReadOnlyList<IrInline> inlines, List<IrTextbox> sink)
    {
        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case IrTextbox tbx: sink.Add(tbx); break;
                case IrFieldRun field: WalkForTextboxes(field.CachedResult, sink); break;
                case IrHyperlink link: WalkForTextboxes(link.Inlines, sink); break;
            }
        }
    }

    // ------------------------------------------------------------------ table diff

    /// <summary>
    /// Verify a nested <see cref="IrTableDiff"/> (M2.2 Task 4): (1) every row/cell anchor resolves
    /// against the ACTUAL left/right tables (row + cell anchors are NOT in the document AnchorIndex, so
    /// we resolve them against the table structure directly); (2) move row pairing is well-formed; (3)
    /// reconstruct the right table's row ContentHash sequence from the left table + the row ops and
    /// assert it equals the actual right table's rows in order; (4) for each ModifyRow, reconstruct each
    /// paired cell's paragraph text from its block ops and assert it equals the right cell's text.
    /// </summary>
    private static void VerifyTableDiff(IrTable left, IrTable right, IrTableDiff diff, IrDiffSettings settings)
    {
        var leftRows = left.Rows.ToDictionary(r => r.Anchor.ToString());
        var rightRows = right.Rows.ToDictionary(r => r.Anchor.ToString());

        // Move-source rows, keyed by group, so a destination can reproduce the moved-from row.
        var moveSourceRow = new Dictionary<int, IrRow>();
        foreach (var op in diff.RowOps)
            if (op.IsMoveSource == true)
                moveSourceRow[op.MoveGroupId!.Value] = ResolveRow(leftRows, op.LeftRowAnchor!);

        var moveSources = new HashSet<int>();
        var moveDests = new HashSet<int>();

        // Reconstruct the right table's rows (ContentHash) in right order.
        var reconstructedRows = new List<IrHash>();
        foreach (var op in diff.RowOps)
        {
            switch (op.Kind)
            {
                case IrRowOpKind.EqualRow:
                {
                    var lr = ResolveRow(leftRows, op.LeftRowAnchor!);
                    var rr = ResolveRow(rightRows, op.RightRowAnchor!);
                    Assert.Equal(lr.ContentHash, rr.ContentHash); // EqualRow ⇒ row content unchanged
                    reconstructedRows.Add(lr.ContentHash);
                    break;
                }
                case IrRowOpKind.ModifyRow:
                {
                    var lr = ResolveRow(leftRows, op.LeftRowAnchor!);
                    var rr = ResolveRow(rightRows, op.RightRowAnchor!);
                    Assert.NotNull(op.CellOps);
                    VerifyCellOps(lr, rr, op.CellOps!, settings);
                    reconstructedRows.Add(rr.ContentHash); // modified content sourced from the right row
                    break;
                }
                case IrRowOpKind.InsertRow:
                {
                    var rr = ResolveRow(rightRows, op.RightRowAnchor!);
                    Assert.Null(op.LeftRowAnchor);
                    reconstructedRows.Add(rr.ContentHash);
                    break;
                }
                case IrRowOpKind.DeleteRow:
                {
                    ResolveRow(leftRows, op.LeftRowAnchor!); // resolves; produces no right row
                    Assert.Null(op.RightRowAnchor);
                    break;
                }
                case IrRowOpKind.MovedRow:
                {
                    Assert.NotNull(op.MoveGroupId);
                    Assert.NotNull(op.IsMoveSource);
                    if (op.IsMoveSource == true)
                    {
                        moveSources.Add(op.MoveGroupId!.Value);
                        ResolveRow(leftRows, op.LeftRowAnchor!);
                    }
                    else
                    {
                        moveDests.Add(op.MoveGroupId!.Value);
                        var rr = ResolveRow(rightRows, op.RightRowAnchor!);
                        var src = moveSourceRow[op.MoveGroupId!.Value];
                        Assert.Equal(src.ContentHash, rr.ContentHash); // exact-content row move
                        reconstructedRows.Add(rr.ContentHash);
                    }
                    break;
                }
            }
        }

        Assert.Equal(moveSources, moveDests); // every moved row has both a source and a destination

        // The reconstructed right-row sequence equals the actual right table's rows in order.
        Assert.Equal(right.Rows.Count, reconstructedRows.Count);
        for (int i = 0; i < right.Rows.Count; i++)
            Assert.Equal(right.Rows[i].ContentHash, reconstructedRows[i]);
    }

    /// <summary>
    /// Verify a ModifyRow's positional cell ops: anchors resolve against the actual rows, paired cells'
    /// block ops reconstruct the right cell text, and the cell-op count covers both rows' cells.
    /// </summary>
    private static void VerifyCellOps(IrRow left, IrRow right, IrNodeList<IrCellOp> cellOps, IrDiffSettings settings)
    {
        var leftCells = left.Cells.ToDictionary(c => c.Anchor.ToString());
        var rightCells = right.Cells.ToDictionary(c => c.Anchor.ToString());

        int leftSeen = 0, rightSeen = 0;
        foreach (var op in cellOps)
        {
            bool hasLeft = op.LeftCellAnchor is not null;
            bool hasRight = op.RightCellAnchor is not null;
            if (hasLeft) { Assert.True(leftCells.ContainsKey(op.LeftCellAnchor!), $"cell anchor {op.LeftCellAnchor} unresolved (left)"); leftSeen++; }
            if (hasRight) { Assert.True(rightCells.ContainsKey(op.RightCellAnchor!), $"cell anchor {op.RightCellAnchor} unresolved (right)"); rightSeen++; }

            if (hasLeft && hasRight && op.BlockOps is { } blockOps)
            {
                // Paired, content-differing cell: reconstruct the right cell's block text from the block
                // ops applied to the left cell's blocks, and assert it matches the right cell's text.
                var lc = leftCells[op.LeftCellAnchor!];
                var rc = rightCells[op.RightCellAnchor!];
                var reconstructed = ReconstructBlocks(lc.Blocks, rc.Blocks, blockOps, settings);
                var expected = BlockTexts(rc.Blocks, settings);
                Assert.Equal(expected, reconstructed);
            }
        }

        Assert.Equal(left.Cells.Count, leftSeen);
        Assert.Equal(right.Cells.Count, rightSeen);
    }

    /// <summary>
    /// Reconstruct a right cell's per-paragraph text (concatenated MatchKeys) from the left cell's blocks
    /// and the cell's block ops, mirroring the body-level reconstruction. Non-paragraph cell blocks
    /// contribute their ContentHash hex as a stand-in text token (compared on both sides identically).
    /// </summary>
    private static List<string> ReconstructBlocks(
        IrNodeList<IrBlock> leftBlocks, IrNodeList<IrBlock> rightBlocks,
        IrNodeList<IrEditOp> blockOps, IrDiffSettings settings)
    {
        var leftByAnchor = leftBlocks.ToDictionary(b => b.Anchor.ToString());
        var rightByAnchor = rightBlocks.ToDictionary(b => b.Anchor.ToString());
        var result = new List<string>();

        foreach (var op in blockOps)
        {
            switch (op.Kind)
            {
                case IrEditOpKind.EqualBlock:
                case IrEditOpKind.FormatOnlyBlock:
                    result.Add(BlockText(leftByAnchor[op.LeftAnchor!], settings));
                    break;
                case IrEditOpKind.ModifyBlock:
                {
                    var lb = leftByAnchor[op.LeftAnchor!];
                    var rb = rightByAnchor[op.RightAnchor!];
                    if (lb is IrParagraph && rb is IrParagraph)
                        result.Add(string.Concat(ApplyModify(lb, rb, op.TokenDiff, settings)!));
                    else
                        result.Add(BlockText(rb, settings)); // nested table/opaque in a cell: source from right
                    break;
                }
                case IrEditOpKind.InsertBlock:
                    result.Add(BlockText(rightByAnchor[op.RightAnchor!], settings));
                    break;
                case IrEditOpKind.DeleteBlock:
                    break;
                case IrEditOpKind.MoveBlock:
                case IrEditOpKind.MoveModifyBlock:
                    // Cell-internal block moves are not produced by the table differ in M2.2; if they
                    // ever are, the destination sources from the right block.
                    if (op.IsMoveSource != true)
                        result.Add(BlockText(rightByAnchor[op.RightAnchor!], settings));
                    break;
            }
        }
        return result;
    }

    private static List<string> BlockTexts(IrNodeList<IrBlock> blocks, IrDiffSettings settings) =>
        blocks.Select(b => BlockText(b, settings)).ToList();

    private static string BlockText(IrBlock block, IrDiffSettings settings) =>
        block is IrParagraph p ? string.Concat(Tokens(p, settings)) : block.ContentHash.ToHex();

    private static IrRow ResolveRow(Dictionary<string, IrRow> rows, string anchor)
    {
        Assert.True(rows.TryGetValue(anchor, out var row), $"row anchor '{anchor}' missing in table.");
        return row!;
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
        MaskTextboxKeys(IrDiffTokenizer.Tokenize(p, settings)).Select(t => t.MatchKey).ToList();

    /// <summary>
    /// The text view masks every Textbox placeholder token's MatchKey to one constant (M2.4 Task 1): a
    /// textbox change is verified through the op's nested <see cref="IrEditOp.TextboxDiffs"/>, and the
    /// builder masks the placeholder out of the paragraph's own token diff, so the paragraph-text
    /// reconstruction must mask it too (the real per-textbox key differs when the interior changed). For an
    /// unchanged textbox the masking is a harmless no-op (both sides already share the same constant).
    /// </summary>
    private const string MaskedTextboxKey = "tbx";

    private static IReadOnlyList<IrDiffToken> MaskTextboxKeys(IReadOnlyList<IrDiffToken> tokens)
    {
        var masked = new List<IrDiffToken>(tokens.Count);
        foreach (var t in tokens)
            masked.Add(t.Kind == IrDiffTokenKind.Textbox ? t with { MatchKey = MaskedTextboxKey } : t);
        return masked;
    }

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

        // Mask textbox placeholder keys so the diff (built over masked tokens when textboxes nest) verifies
        // against the same masked view; harmless when no textbox is present.
        var leftTokens = MaskTextboxKeys(IrDiffTokenizer.Tokenize(lp, settings));
        var rightTokens = MaskTextboxKeys(IrDiffTokenizer.Tokenize(rp, settings));

        // Enforce the full TokenDiff totality/coverage/per-kind battery here too, so the apply-verifier
        // does NOT pass on a structurally-broken-but-text-equal diff (non-tiling spans, wrong-length
        // Equal/FormatChanged runs, corrupt right spans). This makes the concatenated-text check below a
        // SUFFICIENT-given-tiling proof rather than the sole assertion.
        IrTokenDiffAsserts.AssertInvariants(leftTokens, rightTokens, tokenDiff!, settings);

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
