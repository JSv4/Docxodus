using System;
using System.Collections.Generic;
using System.Linq;

namespace Docxodus.Ir.Diff;

/// <summary>Which paired paragraph owns each view of a comment moved to a different paragraph.
/// A comment attached to a paragraph that itself moved follows the paragraph's existing move plan.</summary>
internal sealed class IrCommentAttachments
{
    internal static readonly IrCommentAttachments Empty = new();
    private Dictionary<string, List<(string Id, bool Insert)>>? _byBase;

    internal IReadOnlyList<(string Id, bool Insert)>? For(string? anchor) =>
        anchor is null ? null : _byBase?.GetValueOrDefault(anchor);

    internal static IrCommentAttachments Plan(IrDocument original, IrDocument revised, IReadOnlyList<IrEditOp> operations)
    {
        if (original.Comments.Comments.Count == 0 || revised.Comments.Comments.Count == 0)
            return Empty;

        // The input preparation aligns definition IDs across sources. Only a range wholly in one
        // paragraph at both ends can have two independent attachments of this shape.
        static Dictionary<string, string> Targets(IrDocument document)
        {
            var result = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var comment in document.Comments.Comments)
                if (comment.Targets.Count == 1 && comment.PackageId is { Length: > 0 } id)
                    result.TryAdd(id, comment.Targets[0].BlockAnchor.ToString());
            return result;
        }
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
        var pairedOriginal = pairs.Values.ToHashSet(StringComparer.Ordinal);
        foreach (var (id, oldParent) in originalTargets)
        {
            if (!revisedTargets.TryGetValue(id, out var revisedParent) ||
                !pairs.TryGetValue(revisedParent, out var newParent) || oldParent == newParent || !pairedOriginal.Contains(oldParent))
                continue;
            plan ??= new IrCommentAttachments();
            plan.Add(oldParent, id, false);
            plan.Add(newParent, id, true);
        }
        return plan ?? Empty;
    }

    private void Add(string anchor, string id, bool insert)
    {
        _byBase ??= new(StringComparer.Ordinal);
        if (!_byBase.TryGetValue(anchor, out var list))
            _byBase.Add(anchor, list = new());
        list.Add((id, insert));
    }
}
