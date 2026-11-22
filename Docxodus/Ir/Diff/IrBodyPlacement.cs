using System.Collections.Generic;

namespace Docxodus.Ir.Diff;

/// <summary>A standalone paragraph operation or a gap containing the insertions and deletions
/// between two surviving blocks. Each source retains its original order within the gap.</summary>
internal readonly record struct IrBodyStep(int Start, int End, bool IsGap)
{
    public IEnumerable<IrEditOp> Side(IReadOnlyList<IrEditOp> operations, IrEditOpKind kind)
    {
        for (int i = Start; i < End; i++)
            if (operations[i].Kind == kind)
                yield return operations[i];
    }
}

/// <summary>Physical paragraph order and paragraph-mark ownership shared by story binding and markup.</summary>
internal static class IrBodyPlacement
{
    /// <summary>Whether a modified paragraph occupies two physical paragraphs in the redline.
    /// The public edit script can still describe token and textbox changes inside these copies.</summary>
    internal static bool RequiresCopies(IrEditOp op, IrDocument left, IrDocument right,
        IReadOnlySet<string>? preservedReplacements)
    {
        if (op.RequiresWholeParagraphReplace || op.RightAnchor is { } anchor && preservedReplacements?.Contains(anchor) == true)
            return true;
        if (op.TextboxDiffs is not { } boxes) return false;
        return !CanBind(left, op.LeftAnchor) || !CanBind(right, op.RightAnchor);

        bool CanBind(IrDocument source, string? key) =>
            key is not null && source.AnchorIndex.TryGetValue(key, out var block) && block is IrParagraph paragraph &&
            paragraph.PhysicalTextboxes is { HasBlockMarkers: false, HasSpanningField: false, HasCarrierRelationships: false } facts &&
            facts.Count == boxes.Count;
    }

    internal static IEnumerable<IrBodyStep> Steps(IReadOnlyList<IrEditOp> operations, bool renderMoves)
    {
        bool InGap(IrEditOp op) => op.Kind is IrEditOpKind.InsertBlock or IrEditOpKind.DeleteBlock &&
            !(renderMoves && op.RelocationGroupId.HasValue);

        for (int start = 0; start < operations.Count;)
        {
            int end = start + 1;
            bool gap = InGap(operations[start]);
            if (gap)
                while (end < operations.Count && InGap(operations[end])) end++;
            yield return new IrBodyStep(start, end, gap);
            start = end;
        }
    }

    /// <summary>The final member keeps the paired paragraph mark. Earlier split members gain a
    /// mark, and earlier merge members lose their original mark on acceptance.</summary>
    internal static IrParagraphMember Member(IrEditOp op, int index)
    {
        var members = op.SplitMergeAnchors!;
        bool pairedMark = index == members.Count - 1;
        return op.Kind == IrEditOpKind.SplitBlock
            ? new IrParagraphMember(op.LeftAnchor, members[index], pairedMark, true)
            : new IrParagraphMember(members[index], pairedMark ? op.RightAnchor : null, true, pairedMark);
    }
}

internal readonly record struct IrParagraphMember(string? Left, string? Right, bool LeftMark, bool RightMark);
