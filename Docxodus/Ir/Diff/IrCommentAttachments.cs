using System;
using System.Collections.Generic;
using System.Linq;

namespace Docxodus.Ir.Diff;

/// <summary>Which source paragraph owns each view of a comment moved to a different paragraph.
/// A comment attached to a paragraph that itself moved follows the paragraph's existing move plan.</summary>
internal sealed class IrCommentAttachments
{
    internal enum View { Original, Revised, Both }

    internal static readonly IrCommentAttachments Empty = new();
    private Dictionary<string, List<(string Id, View View, string? Author)>>? _byBase;

    private Dictionary<string, string>? _anchorByUnid;
    private HashSet<IrAnchor>? _unchanged;
    private HashSet<(int Reviewer, string Id)>? _selectedReattachments;

    internal bool HasPlacements => _byBase is not null;
    internal string? AnchorForCopy(string? unid) => unid is null ? null : _anchorByUnid?.GetValueOrDefault(unid);

    internal IReadOnlyList<(string Id, View View, string? Author)>? For(string? anchor) =>
        anchor is null ? null : _byBase?.GetValueOrDefault(anchor);

    internal static IrCommentAttachments Plan(IrDocument original, IrDocument revised, IReadOnlyList<IrEditOp> operations)
    {
        if (original.Comments.Comments.Count == 0) return Empty;
        var originalTargets = Targets(original);
        var revisedTargets = Targets(revised);
        var baseEndpoints = new HashSet<IrAnchor>();
        foreach (var comment in original.Comments.Comments)
            if (comment.BodyMarkerParagraphs is { } range)
            {
                baseEndpoints.Add(range.First);
                baseEndpoints.Add(range.Last);
            }
        if (baseEndpoints.Count == 0) return Empty;
        var revisedEndpoints = revisedTargets.Values.ToHashSet();
        var pairs = new Dictionary<IrAnchor, IrAnchor>();
        var sources = new HashSet<IrAnchor>();
        var plan = new IrCommentAttachments();
        void Pair(IrParagraph left, IrParagraph? right, bool kept, bool reattach)
        {
            if (baseEndpoints.Contains(left.Anchor))
            {
                if (kept) (plan._unchanged ??= new()).Add(left.Anchor);
                if (reattach) sources.Add(left.Anchor);
            }
            if (reattach && right is not null && revisedEndpoints.Contains(right.Anchor))
                pairs[right.Anchor] = left.Anchor;
        }

        // Equal containers have the same block order, including nested cells and controls.
        // Edited containers name each matched row, cell and paragraph in their child operations.
        void Equal(IEnumerable<IrBlock> left, IEnumerable<IrBlock> right, bool reattach = true)
        {
            foreach (var (l, r) in left.SelectMany(Paragraphs).Zip(right.SelectMany(Paragraphs)))
                Pair(l, r, true, reattach);
        }
        void Walk(IReadOnlyList<IrEditOp> ops)
        {
            foreach (var op in ops)
            {
                var left = op.LeftAnchor is { } l ? original.AnchorIndex.GetValueOrDefault(l) : null;
                var right = op.RightAnchor is { } r ? revised.AnchorIndex.GetValueOrDefault(r) : null;
                if (left is null) continue;
                if (op.Kind == IrEditOpKind.DeleteBlock && op.RelocationGroupId is null)
                    foreach (var paragraph in Paragraphs(left)) Pair(paragraph, null, false, true);
                else if (right is not null && op.Kind is IrEditOpKind.EqualBlock or IrEditOpKind.FormatOnlyBlock or IrEditOpKind.MoveBlock)
                {
                    if (left is IrParagraph lpara && right is IrParagraph rpara)
                        Pair(lpara, rpara, true, op.Kind != IrEditOpKind.MoveBlock);
                    else
                        Equal(new[] { left }, new[] { right }, op.Kind != IrEditOpKind.MoveBlock);
                }
                else if (op.Kind == IrEditOpKind.ModifyBlock && left is IrParagraph lp && right is IrParagraph rp)
                    Pair(lp, rp, op.TokenDiff is { } tokens && !tokens.Ops.Any(t => t.Kind is IrTokenOpKind.Insert or IrTokenOpKind.Delete), true);
                if (op.WrapperOps is { } children) Walk(children);
                if (op.TextboxDiffs is { } boxes)
                    foreach (var box in boxes) Walk(box.Ops);
                if (op.TableDiff is not { } table) continue;
                foreach (var row in table.RowOps)
                {
                    if (row.Kind == IrRowOpKind.DeleteRow && row.LeftRowAnchor is { } deletedRow &&
                        IrCompositeMerger.RowIndex(original).TryGetValue(deletedRow, out var removed))
                        foreach (var paragraph in removed.Cells.SelectMany(c => c.Blocks).SelectMany(Paragraphs))
                            Pair(paragraph, null, false, true);
                    if (row.Kind == IrRowOpKind.EqualRow && row.LeftRowAnchor is { } lr && row.RightRowAnchor is { } rr &&
                        IrCompositeMerger.RowIndex(original).TryGetValue(lr, out var leftRow) &&
                        IrCompositeMerger.RowIndex(revised).TryGetValue(rr, out var rightRow))
                        Equal(leftRow.Cells.SelectMany(c => c.Blocks), rightRow.Cells.SelectMany(c => c.Blocks));
                    if (row.CellOps is not { } cells) continue;
                    foreach (var cell in cells)
                        if (cell.BlockOps is { } blocks) Walk(blocks);
                        else if (cell.RightCellAnchor is null && cell.LeftCellAnchor is { } deletedCell &&
                            IrCompositeMerger.CellIndex(original).TryGetValue(deletedCell, out var removedCell))
                        {
                            foreach (var paragraph in removedCell.Blocks.SelectMany(Paragraphs))
                                Pair(paragraph, null, false, true);
                        }
                        else if (cell.LeftCellAnchor is { } lc && cell.RightCellAnchor is { } rc &&
                            IrCompositeMerger.CellIndex(original).TryGetValue(lc, out var leftCell) &&
                            IrCompositeMerger.CellIndex(revised).TryGetValue(rc, out var rightCell))
                            Equal(leftCell.Blocks, rightCell.Blocks);
                }
            }
        }
        Walk(operations);
        foreach (var (id, oldParent) in originalTargets)
        {
            if (!revisedTargets.TryGetValue(id, out var revisedParent) ||
                !pairs.TryGetValue(revisedParent, out var newParent) || oldParent == newParent || !sources.Contains(oldParent))
                continue;
            plan.Add(oldParent, id, View.Original);
            plan.Add(newParent, id, View.Revised);
        }
        return plan;
    }

    private static IEnumerable<IrParagraph> Paragraphs(IrBlock block) => block switch
    {
        IrParagraph p => new[] { p },
        IrSdtBlock wrapper => wrapper.Blocks.SelectMany(Paragraphs),
        IrTable table => table.Rows.SelectMany(row => row.Cells).SelectMany(cell => cell.Blocks).SelectMany(Paragraphs),
        _ => Enumerable.Empty<IrParagraph>()
    };

    private static Dictionary<string, IrAnchor> Targets(IrDocument document)
    {
        var result = new Dictionary<string, IrAnchor>(StringComparer.Ordinal);
        foreach (var comment in document.Comments.Comments)
            if (comment.BodyMarkerParagraphs is { } range && range.First == range.Last && comment.PackageId is { Length: > 0 } id)
                result.TryAdd(id, range.First);
        return result;
    }

    /// <summary>A reviewer can reattach a comment after deleting or rewriting its original
    /// paragraph when its new target is a surviving base copy. An edited target's owner
    /// supplies its comments. A reviewer who only moves an existing base comment leaves it with the base.</summary>
    internal static IrCommentAttachments? Compose(IrDocument original, IReadOnlyList<IrEditScript> comparisons,
        IReadOnlyList<(string Author, IrDocument Ir)> reviewers, IReadOnlyList<IrCompositeOp> operations, ConflictResolution policy, List<IrConflict> conflicts, ref int nextConflictId)
    {
        if (comparisons.Count == 1) return comparisons[0].CommentAttachments;
        if (!comparisons.Any(c => c.CommentAttachments?._byBase is not null)) return null;
        var originalTargets = Targets(original);
        var owners = new Dictionary<string, IrCompositeOp>(StringComparer.Ordinal);
        void Index(IrCompositeOp op)
        {
            if (op.Op.LeftAnchor is not { } anchor || !original.AnchorIndex.TryGetValue(anchor, out var block)) return;
            if (op.Op.Kind is IrEditOpKind.EqualBlock or IrEditOpKind.DeleteBlock)
            {
                foreach (var paragraph in Paragraphs(block)) owners.TryAdd(paragraph.Anchor.ToString(), op);
                return;
            }
            if (block is IrParagraph) owners.TryAdd(anchor, op);
            if (op.AuthoredWrapperOps is { } children)
                foreach (var child in children) Index(child);
            else if (op.Op.WrapperOps is { } wrapper)
                foreach (var child in wrapper) Index(new IrCompositeOp(child, op.Author, op.SourceReviewer));
            if (op.AuthoredRows is not { } rows) return;
            var equal = op with { Op = op.Op with { Kind = IrEditOpKind.EqualBlock } };
            foreach (var row in rows)
            {
                if (row.Kind is IrRowOpKind.EqualRow or IrRowOpKind.DeleteRow && row.BaseRowAnchor is { } rowId &&
                    IrCompositeMerger.RowIndex(original).TryGetValue(rowId, out var source))
                    foreach (var paragraph in source.Cells.SelectMany(c => c.Blocks).SelectMany(Paragraphs))
                        owners.TryAdd(paragraph.Anchor.ToString(), row.Kind == IrRowOpKind.EqualRow ? equal
                            : op with { Op = op.Op with { Kind = IrEditOpKind.DeleteBlock }, SourceReviewer = row.SourceReviewer });
                if (row.ComposedCells is not { } cells) continue;
                foreach (var cell in cells)
                    if (cell.ComposedBlockOps is { } ops)
                        foreach (var child in ops) Index(child);
                    else if (cell.Kind is IrAuthoredCellKind.Content or IrAuthoredCellKind.DeleteCell && cell.BaseCellAnchor is { } cellId &&
                        IrCompositeMerger.CellIndex(original).TryGetValue(cellId, out var sourceCell))
                        foreach (var paragraph in sourceCell.Blocks.SelectMany(Paragraphs))
                            owners.TryAdd(paragraph.Anchor.ToString(), cell.Kind == IrAuthoredCellKind.Content ? equal
                                : op with { Op = op.Op with { Kind = IrEditOpKind.DeleteBlock }, SourceReviewer = cell.ShellSourceReviewer });
            }
        }
        foreach (var op in operations) Index(op);
        var candidates = new Dictionary<string, List<(int Reviewer, IrAnchor Origin, IrAnchor Destination)>>(StringComparer.Ordinal);
        for (int r = 0; r < comparisons.Count; r++)
        {
            if (comparisons[r].CommentAttachments?._byBase is not { } placements) continue;
            foreach (var (destination, entries) in placements)
            {
                if (!owners.TryGetValue(destination, out var target) ||
                    target.Op.Kind != IrEditOpKind.EqualBlock && !(target.Op.Kind == IrEditOpKind.ModifyBlock && Contributes(target, r))) continue;
                foreach (var (id, view, _) in entries)
                {
                    if (view != View.Revised || !originalTargets.TryGetValue(id, out var origin) || !owners.TryGetValue(origin.ToString(), out var source) ||
                        source.Op.Kind is not (IrEditOpKind.DeleteBlock or IrEditOpKind.ModifyBlock)) continue;
                    if (!Contributes(source, r)) continue;
                    if (!candidates.TryGetValue(id, out var entriesForId)) candidates.Add(id, entriesForId = new());
                    entriesForId.Add((r, origin, original.AnchorIndex[destination].Anchor));
                }
            }
        }
        IrCommentAttachments? result = null;
        foreach (var (id, choices) in candidates)
        {
            result ??= new IrCommentAttachments();
            var first = choices[0];
            if (choices.Any(c => c.Destination != first.Destination))
            {
                // One logical comment has one accepted range. Resolve competing destinations
                // like other single-valued properties in a composite document.
                conflicts.Add(new IrConflict(nextConflictId++, first.Origin.ToString(), 0, 0, policy,
                    IrNodeList.From(choices.Select(c => new IrConflictCompetitor(reviewers[c.Reviewer].Author, c.Destination.ToString())))));
                if (policy == ConflictResolution.BaseWins)
                {
                    result.Add(first.Origin, id, View.Both);
                    continue;
                }
            }
            (result._selectedReattachments ??= new()).Add((first.Reviewer, id));
            result.Add(first.Origin, id, View.Original, reviewers[first.Reviewer].Author);
            result.Add(first.Destination, id, View.Revised, reviewers[first.Reviewer].Author);
        }
        return result;
    }

    private static bool Contributes(IrCompositeOp source, int reviewer) => source.AuthoredTokens is { } tokens
        ? tokens.Any(t => t.SourceReviewer == reviewer && t.Op.Kind != IrTokenOpKind.Equal)
        : source.SourceReviewer == reviewer;

    /// <summary>For each base body comment, decide whether this reviewer kept the original
    /// anchor paragraphs. An unchanged paragraph elsewhere with identical text has its own anchor.</summary>
    internal static Dictionary<string, bool>? BaseAttachments(IrDocument original, IrEditScript comparison,
        bool soleReviewer, int reviewer, IrCommentAttachments? selected)
    {
        var comments = original.Comments.Comments;
        if (comments.Count == 0) return null;
        var unchanged = comparison.CommentAttachments?._unchanged;
        Dictionary<string, bool>? result = null;
        foreach (var comment in comments)
        {
            if (comment.PackageId is not { } id || comment.BodyMarkerParagraphs is not { } range ||
                range.First.Scope != "body") continue;
            var start = range.First.ToString();
            bool reattached = comparison.CommentAttachments?.For(start)?.Any(p => p.Id == id && p.View == View.Original) == true;
            if (soleReviewer && reattached) continue;
            bool suppressed = reattached && selected?._selectedReattachments?.Contains((reviewer, id)) != true;
            (result ??= new(StringComparer.Ordinal))[id] =
                suppressed || unchanged?.Contains(range.First) == true && unchanged.Contains(range.Last);
        }
        return result;
    }

    private void Add(IrAnchor paragraph, string id, View view, string? author = null)
    {
        var anchor = paragraph.ToString();
        (_anchorByUnid ??= new(StringComparer.Ordinal))[paragraph.Unid] = anchor;
        _byBase ??= new(StringComparer.Ordinal);
        if (!_byBase.TryGetValue(anchor, out var list))
            _byBase.Add(anchor, list = new());
        list.Add((id, view, author));
    }
}
