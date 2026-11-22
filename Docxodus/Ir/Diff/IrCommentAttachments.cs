using System;
using System.Collections.Generic;
using System.Linq;

namespace Docxodus.Ir.Diff;

/// <summary>Which source paragraph owns each view of a comment moved to a different paragraph.
/// A comment attached to a paragraph that itself moved follows the paragraph's existing move plan.</summary>
internal sealed class IrCommentAttachments
{
    internal static readonly IrCommentAttachments Empty = new();
    private Dictionary<string, List<(string Id, bool Insert, string? Author)>>? _byBase;

    internal IReadOnlyList<(string Id, bool Insert, string? Author)>? For(string? anchor) =>
        anchor is null ? null : _byBase?.GetValueOrDefault(anchor);

    internal static IrCommentAttachments Plan(IrDocument original, IrDocument revised, IReadOnlyList<IrEditOp> operations)
    {
        if (original.Comments.Comments.Count == 0 || revised.Comments.Comments.Count == 0)
            return Empty;

        // The input preparation aligns definition IDs across sources. Only a range wholly in one
        // paragraph at both ends can have two independent attachments of this shape.
        var originalTargets = Targets(original);
        var revisedTargets = Targets(revised);
        if (!originalTargets.Keys.Any(revisedTargets.ContainsKey))
            return Empty;
        var pairs = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var op in operations)
            if (op.Kind is IrEditOpKind.EqualBlock or IrEditOpKind.FormatOnlyBlock or IrEditOpKind.ModifyBlock &&
                op.LeftAnchor is { } left && op.RightAnchor is { } right)
                pairs[right] = left;

        IrCommentAttachments? plan = null;
        var originalSources = pairs.Values.ToHashSet(StringComparer.Ordinal);
        foreach (var op in operations)
            if (op.Kind == IrEditOpKind.DeleteBlock && op.RelocationGroupId is null && op.LeftAnchor is { } deleted)
                originalSources.Add(deleted);
        foreach (var (id, oldParent) in originalTargets)
        {
            if (!revisedTargets.TryGetValue(id, out var revisedParent) ||
                !pairs.TryGetValue(revisedParent, out var newParent) || oldParent == newParent || !originalSources.Contains(oldParent))
                continue;
            plan ??= new IrCommentAttachments();
            plan.Add(oldParent, id, false);
            plan.Add(newParent, id, true);
        }
        return plan ?? Empty;
    }

    private static Dictionary<string, string> Targets(IrDocument document)
    {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var comment in document.Comments.Comments)
            if (comment.Targets.Count == 1 && comment.PackageId is { Length: > 0 } id)
                result.TryAdd(id, comment.Targets[0].BlockAnchor.ToString());
        return result;
    }

    /// <summary>A reviewer can reattach a comment after deleting or rewriting its original
    /// paragraph when its new target is a surviving base copy. An edited target's owner
    /// supplies its comments. A reviewer who only moves an existing base comment leaves it with the base.</summary>
    internal static IrCommentAttachments? Compose(IrDocument original, IReadOnlyList<IrEditScript> comparisons,
        IReadOnlyList<(string Author, IrDocument Ir)> reviewers, IReadOnlyList<IrCompositeOp> operations)
    {
        if (comparisons.Count == 1) return comparisons[0].CommentAttachments;
        if (!comparisons.Any(c => c.CommentAttachments?._byBase is not null)) return null;
        var originalTargets = Targets(original);
        var owners = new Dictionary<string, IrCompositeOp>(StringComparer.Ordinal);
        foreach (var op in operations)
            if (op.Op.LeftAnchor is { } anchor) owners.TryAdd(anchor, op);
        IrCommentAttachments? result = null;
        for (int r = 0; r < comparisons.Count; r++)
        {
            if (comparisons[r].CommentAttachments?._byBase is not { } placements) continue;
            foreach (var (destination, entries) in placements)
            {
                if (!owners.TryGetValue(destination, out var target) || target.Op.Kind != IrEditOpKind.EqualBlock) continue;
                foreach (var (id, insert, _) in entries)
                {
                    if (!insert || !originalTargets.TryGetValue(id, out var origin) || !owners.TryGetValue(origin, out var source) ||
                        source.SourceReviewer != r || source.Op.Kind is not (IrEditOpKind.DeleteBlock or IrEditOpKind.ModifyBlock)) continue;
                    result ??= new IrCommentAttachments();
                    result.Add(origin, id, false, reviewers[r].Author);
                    result.Add(destination, id, true, reviewers[r].Author);
                }
            }
        }
        return result;
    }

    /// <summary>For each base body comment, decide whether this reviewer kept the original
    /// anchor paragraphs. An unchanged paragraph elsewhere with identical text has its own anchor.</summary>
    internal static Dictionary<string, bool>? BaseAttachments(IrDocument original, IrEditScript comparison,
        bool soleReviewer)
    {
        var comments = original.Comments.Comments;
        if (comments.Count == 0) return null;
        var unchanged = new HashSet<string>(StringComparer.Ordinal);
        void Keep(IrBlock block)
        {
            if (block is IrParagraph paragraph)
                unchanged.Add(paragraph.Anchor.ToString());
            else if (block is IrSdtBlock wrapper)
                foreach (var child in wrapper.Blocks) Keep(child);
            else if (block is IrTable table)
                foreach (var row in table.Rows) KeepCells(row);
        }
        void KeepCells(IrRow row)
        {
            foreach (var cell in row.Cells)
                foreach (var child in cell.Blocks) Keep(child);
        }
        void Walk(IReadOnlyList<IrEditOp> ops)
        {
            foreach (var op in ops)
            {
                if (op.LeftAnchor is { } a && original.AnchorIndex.TryGetValue(a, out var block) &&
                    (op.Kind is IrEditOpKind.EqualBlock or IrEditOpKind.FormatOnlyBlock or IrEditOpKind.MoveBlock ||
                     block is IrParagraph && op.TokenDiff is { } tokens &&
                     !tokens.Ops.Any(t => t.Kind is IrTokenOpKind.Insert or IrTokenOpKind.Delete)))
                    Keep(block);
                if (op.WrapperOps is { } children) Walk(children);
                if (op.TextboxDiffs is { } boxes)
                    foreach (var box in boxes) Walk(box.Ops);
                if (op.TableDiff is not { } table) continue;
                foreach (var row in table.RowOps)
                {
                    if (row.Kind == IrRowOpKind.EqualRow && row.LeftRowAnchor is { } aRow &&
                        IrCompositeMerger.RowIndex(original).TryGetValue(aRow, out var source)) KeepCells(source);
                    if (row.CellOps is not { } cells) continue;
                    foreach (var cell in cells)
                        if (cell.BlockOps is { } blocks) Walk(blocks);
                        else if (cell.LeftCellAnchor is { } aCell && cell.RightCellAnchor is not null &&
                            IrCompositeMerger.CellIndex(original).TryGetValue(aCell, out var originalCell))
                            foreach (var child in originalCell.Blocks) Keep(child);
                }
            }
        }
        Walk(comparison.Operations);
        Dictionary<string, bool>? result = null;
        foreach (var comment in comments)
        {
            if (comment.PackageId is not { } id || comment.Targets.Count == 0 ||
                comment.Targets[0].BlockAnchor.Scope != "body") continue;
            var start = comment.Targets[0].BlockAnchor.ToString();
            if (soleReviewer && comparison.CommentAttachments?.For(start)?.Any(p => p.Id == id && !p.Insert) == true)
                continue; // both copies are explicit in the sole reviewer's attachment plan
            (result ??= new(StringComparer.Ordinal))[id] =
                unchanged.Contains(start) && unchanged.Contains(comment.Targets[^1].BlockAnchor.ToString());
        }
        return result;
    }

    private void Add(string anchor, string id, bool insert, string? author = null)
    {
        _byBase ??= new(StringComparer.Ordinal);
        if (!_byBase.TryGetValue(anchor, out var list))
            _byBase.Add(anchor, list = new());
        list.Add((id, insert, author));
    }
}
