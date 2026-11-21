using System;
using System.Collections.Generic;
using System.Linq;

namespace Docxodus.Ir.Diff;

/// <summary>A control instance in one of the two input containers. Entry indexes address the reader's
/// ordered children, including when the edit plan and the writer use different source snapshots.</summary>
internal readonly record struct IrTableControlId(int Source, int Entry)
{
    public bool Left => Source < 0;
}

internal sealed record IrControlLayout(IrTableControlId[][] Chains, HashSet<IrTableControlId> OneSided,
    Dictionary<int, int>? BaseForRight = null, Dictionary<IrTableControlId, string>? Authors = null);

internal readonly record struct IrPlannedTableRow(int OpIndex, int Left, int Right, IrControlLayout? Cells = null);

internal sealed record IrTableControlPlan(IrPlannedTableRow[] Rows, IrControlLayout? Controls,
    IReadOnlyDictionary<int, int> EmptyControlPositions);

/// <summary>Pairs control instances by the members the edit plan pairs, and places each output
/// row or cell in its original, shared, or newly inserted controls.</summary>
internal static class IrTableControlPlanner
{
    internal static IrTableControlPlan? Plan(IrTable left, IrTable right, IReadOnlyList<IrRowOp> ops)
    {
        if (!IrTableDiffer.ControlsDrawable(left.Controls, right.Controls))
            return null;
        var leftRows = Index(left);
        var rightRows = Index(right);
        var rows = new IrPlannedTableRow[ops.Count];
        for (int i = 0; i < rows.Length; i++)
        {
            var op = ops[i];
            int l = op.LeftRowAnchor is { } la && (op.Kind != IrRowOpKind.MovedRow || op.IsMoveSource == true) ? leftRows.GetValueOrDefault(la, -1) : -1;
            int r = op.RightRowAnchor is { } ra && (op.Kind != IrRowOpKind.MovedRow || op.IsMoveSource != true) ? rightRows.GetValueOrDefault(ra, -1) : -1;
            IrControlLayout? cells = null;
            if (op.Kind == IrRowOpKind.ModifyRow && op.CellOps is { } cellOps && l >= 0 && r >= 0)
            {
                var a = left.Rows[l];
                var b = right.Rows[r];
                if (!IrTableDiffer.ControlsDrawable(a.Controls, b.Controls))
                    return null;
                if (((a.Controls | b.Controls) & IrContainerControls.Control) != 0)
                {
                    int li = 0, ri = 0;
                    var members = new (int Left, int Right)[cellOps.Count];
                    for (int j = 0; j < members.Length; j++)
                        members[j] = (cellOps[j].LeftCellAnchor is null ? -1 : li++, cellOps[j].RightCellAnchor is null ? -1 : ri++);
                    if (li != a.Cells.Count || ri != b.Cells.Count ||
                        PairControls(members, a.ChildOrder, b.ChildOrder) is not { } layout || !StayTogether(layout.Chains))
                        return null;
                    cells = layout;
                }
            }
            rows[i] = new IrPlannedTableRow(i, l, r, cells);
        }

        IReadOnlyList<IrRowOp> drawn = ops;
        IrControlLayout? controls = null;
        if (((left.Controls | right.Controls) & IrContainerControls.Control) != 0)
        {
            controls = PairControls(rows.Select(r => (r.Left, r.Right)).ToArray(), left.ChildOrder, right.ChildOrder);
            if (controls is null)
                return null;
            if (!StayTogether(controls.Chains))
            {
                var order = KeepControlsTogether(ops, controls.Chains);
                controls = controls with { Chains = order.Select(i => controls.Chains[i]).ToArray() };
                if (!StayTogether(controls.Chains))
                    return null;
                rows = order.Select(i => rows[i]).ToArray();
                drawn = order.Select(i => ops[i]).ToArray();
            }
        }
        return IrTableDiffer.ControlsInPlace(left, right, drawn) is { } places
            ? new IrTableControlPlan(rows, controls, places) : null;
    }

    private static Dictionary<string, int> Index(IrTable table)
    {
        var index = new Dictionary<string, int>(table.Rows.Count, StringComparer.Ordinal);
        for (int i = 0; i < table.Rows.Count; i++)
            index[table.Rows[i].Anchor.ToString()] = i;
        return index;
    }

    private static IrTableControlId[] Chain(IrTableOrder? order, int member, int source)
    {
        if (order is null || member < 0)
            return Array.Empty<IrTableControlId>();
        int parent = order.Children[order.Members[member]].Parent;
        if (parent < 0)
            return Array.Empty<IrTableControlId>();
        var chain = new List<IrTableControlId>();
        for (int index = parent; index >= 0; index = order.Children[index].Parent)
            chain.Add(new IrTableControlId(source, index));
        chain.Reverse();
        return chain.ToArray();
    }

    private static IrControlLayout? PairControls(IReadOnlyList<(int Left, int Right)> members,
        IrTableOrder? left, IrTableOrder? right)
    {
        var leftToRight = new Dictionary<IrTableControlId, IrTableControlId>();
        var pairedRight = new HashSet<IrTableControlId>();
        var chains = new IrTableControlId[members.Count][];
        for (int i = 0; i < chains.Length; i++)
        {
            var (l, r) = members[i];
            var rightChain = Chain(right, r, 0);
            var leftChain = Chain(left, l, -1);
            chains[i] = r >= 0 ? rightChain : leftChain;
            if (l < 0 || r < 0)
                continue;
            if (leftChain.Length != rightChain.Length)
                return null;
            for (int level = 0; level < leftChain.Length; level++)
            {
                var a = leftChain[level];
                var b = rightChain[level];
                if (leftToRight.TryGetValue(a, out var existing))
                {
                    if (existing != b) return null;
                }
                else
                {
                    if (!pairedRight.Add(b) || left!.Children[a.Entry].Digest != right!.Children[b.Entry].Digest)
                        return null;
                    leftToRight.Add(a, b);
                }
            }
        }
        var oneSided = new HashSet<IrTableControlId>();
        foreach (var chain in chains)
            for (int i = 0; i < chain.Length; i++)
            {
                if (leftToRight.TryGetValue(chain[i], out var shared))
                    chain[i] = shared;
                else if (!pairedRight.Contains(chain[i]))
                    oneSided.Add(chain[i]);
            }
        return new IrControlLayout(chains, oneSided, leftToRight.ToDictionary(p => p.Value.Entry, p => p.Key.Entry));
    }

    internal static bool StayTogether(IrTableControlId[][] chains)
    {
        var closed = new HashSet<IrTableControlId>();
        var previous = Array.Empty<IrTableControlId>();
        foreach (var chain in chains)
        {
            int common = Common(previous, chain);
            for (int i = common; i < previous.Length; i++) closed.Add(previous[i]);
            for (int i = common; i < chain.Length; i++)
                if (closed.Contains(chain[i])) return false;
            previous = chain;
        }
        return true;
    }

    private static int Common(IrTableControlId[] a, IrTableControlId[] b)
    {
        int k = 0;
        while (k < a.Length && k < b.Length && a[k] == b[k]) k++;
        return k;
    }

    /// <summary>
    /// An order for <paramref name="ops"/> that keeps each control's rows together: each gap (a run of deleted and
    /// inserted rows between rows both tables keep) is merged, deleted rows in their order and inserted rows in theirs,
    /// taking next the row whose controls go on from the row before it, then the one that shares less with the row after
    /// the gap, then a deleted row.
    /// </summary>
    private static List<int> KeepControlsTogether(IReadOnlyList<IrRowOp> ops, IrTableControlId[][] chains)
    {
        static bool InGap(IrRowOp op) => op.Kind is IrRowOpKind.DeleteRow or IrRowOpKind.InsertRow;
        var order = new List<int>(ops.Count);
        for (int i = 0; i < ops.Count;)
        {
            if (!InGap(ops[i]))
            {
                order.Add(i++);
                continue;
            }
            int end = i;
            while (end < ops.Count && InGap(ops[end]))
                end++;
            var deletes = new List<int>();
            var inserts = new List<int>();
            for (int k = i; k < end; k++)
                (ops[k].Kind == IrRowOpKind.DeleteRow ? deletes : inserts).Add(k);
            var open = order.Count > 0 ? chains[order[^1]] : Array.Empty<IrTableControlId>();
            var next = end < ops.Count ? chains[end] : Array.Empty<IrTableControlId>();
            int d = 0, n = 0;
            while (d < deletes.Count || n < inserts.Count)
            {
                bool delete = n == inserts.Count || d < deletes.Count && (
                    Common(open, chains[deletes[d]]) != Common(open, chains[inserts[n]])
                        ? Common(open, chains[deletes[d]]) > Common(open, chains[inserts[n]])
                        : Common(next, chains[deletes[d]]) <= Common(next, chains[inserts[n]]));
                int pick = delete ? deletes[d++] : inserts[n++];
                order.Add(pick);
                open = chains[pick];
            }
            i = end;
        }
        return order;
    }


    /// <summary>Project reviewer insertions into the base control instances paired by their own
    /// comparisons. An unpaired inserted control belongs to its reviewer. A base control whose
    /// last member was deleted belongs to the first reviewer who deleted a member.</summary>
    internal static IrComposedTableControls? Compose(IrTable original, List<IrAuthoredRowOp> rows,
        IReadOnlyList<IrReviewerTable> contributions, IReadOnlyDictionary<int, IrTable> tables)
    {
        var baseRows = Index(original);
        var reviewRows = tables.ToDictionary(t => t.Key, t => Index(t.Value));
        var baseForReviewer = new Dictionary<int, Dictionary<int, int>>();
        foreach (var contribution in contributions)
        {
            var reviewer = contribution.Reviewer;
            var table = tables[reviewer];
            if (((original.Controls | table.Controls) & IrContainerControls.Control) != 0)
            {
                // A marker-only reviewer can bring a synthetic all-equal row script.
                var comparison = contribution.Diff.ControlPlan ?? Plan(original, table, contribution.Diff.RowOps);
                if (comparison?.Controls?.BaseForRight is not { } pairs) return null;
                baseForReviewer[reviewer] = pairs;
            }
        }
        var chains = new IrTableControlId[rows.Count][];
        var authors = new Dictionary<IrTableControlId, string>();
        var surviving = new HashSet<IrTableControlId>();
        var cells = new Dictionary<string, (int Row, IrControlLayout Layout)>(StringComparer.Ordinal);
        for (int i = 0; i < rows.Count; i++)
        {
            var row = rows[i];
            if (row.Kind == IrRowOpKind.InsertRow)
            {
                int reviewer = row.SourceReviewer;
                if (!tables.TryGetValue(reviewer, out var table) || row.RightRowAnchor is not { } anchor ||
                    !reviewRows[reviewer].TryGetValue(anchor, out var ordinal)) return null;
                var chain = Chain(table.ChildOrder, ordinal, reviewer);
                for (int j = 0; j < chain.Length; j++)
                    if (baseForReviewer.TryGetValue(reviewer, out var pairs) && pairs.TryGetValue(chain[j].Entry, out int shared))
                    {
                        chain[j] = new IrTableControlId(-1, shared);
                        surviving.Add(chain[j]);
                    }
                    else authors[chain[j]] = row.Author;
                chains[i] = chain;
                continue;
            }
            if (row.BaseRowAnchor is not { } originalAnchor || !baseRows.TryGetValue(originalAnchor, out int baseOrdinal))
                return null;
            chains[i] = Chain(original.ChildOrder, baseOrdinal, -1);
            foreach (var control in chains[i])
            {
                if (row.Kind == IrRowOpKind.DeleteRow) authors.TryAdd(control, row.Author);
                else surviving.Add(control);
            }
            var source = original.Rows[baseOrdinal];
            if (row.ComposedCells is not { } outputCells || (source.Controls & IrContainerControls.Control) == 0)
                continue;
            if (!IrTableDiffer.ControlsDrawable(source.Controls, source.Controls) ||
                outputCells.Count != source.Cells.Count || outputCells.Any(c => c.Kind != IrAuthoredCellKind.Content))
                return null;
            foreach (var contribution in contributions)
                foreach (var op in contribution.Diff.RowOps)
                    if (op.LeftRowAnchor == originalAnchor && op.RightRowAnchor is { } right &&
                        reviewRows[contribution.Reviewer].TryGetValue(right, out var reviewedOrdinal) &&
                        !IrTableDiffer.ControlsDrawable(source.Controls, tables[contribution.Reviewer].Rows[reviewedOrdinal].Controls))
                        return null;
            var cellChains = new IrTableControlId[outputCells.Count][];
            for (int j = 0; j < outputCells.Count; j++)
            {
                if (outputCells[j].BaseCellAnchor != source.Cells[j].Anchor.ToString()) return null;
                cellChains[j] = Chain(source.ChildOrder, j, -1);
            }
            cells[originalAnchor] = (baseOrdinal, new IrControlLayout(cellChains, new()));
        }
        foreach (var shared in surviving) authors.Remove(shared);
        List<int>? order = null;
        if (!StayTogether(chains))
        {
            order = KeepInsertsTogether(rows, chains);
            chains = order.Select(i => chains[i]).ToArray();
            if (!StayTogether(chains)) return null;
        }
        var layout = chains.Any(c => c.Length > 0) ? new IrControlLayout(chains, authors.Keys.ToHashSet(), Authors: authors) : null;
        return new IrComposedTableControls(layout, cells, order, tables.ToDictionary(t => t.Key, t => t.Value.Anchor.ToString()));
    }

    /// <summary>
    /// An order for a composed table's rows that keeps each control's rows together: each run of rows that reviewers
    /// insert in one place is merged, each reviewer's in that reviewer's order, taking next the row whose controls go on
    /// from the row before it, then the one that shares less with the row after the run, then the earlier reviewer's.
    /// </summary>
    private static List<int> KeepInsertsTogether(IReadOnlyList<IrAuthoredRowOp> rows, IrTableControlId[][] chains)
    {
        var order = new List<int>(rows.Count);
        for (int i = 0; i < rows.Count;)
        {
            if (rows[i].Kind != IrRowOpKind.InsertRow)
            {
                order.Add(i++);
                continue;
            }
            int end = i;
            while (end < rows.Count && rows[end].Kind == IrRowOpKind.InsertRow)
                end++;
            var queues = Enumerable.Range(i, end - i).GroupBy(k => rows[k].SourceReviewer).OrderBy(g => g.Key)
                .Select(g => new Queue<int>(g)).ToList();
            var open = order.Count > 0 ? chains[order[^1]] : Array.Empty<IrTableControlId>();
            var next = end < rows.Count ? chains[end] : Array.Empty<IrTableControlId>();
            while (queues.Any(q => q.Count > 0))
            {
                var pick = queues.Where(q => q.Count > 0)
                    .OrderByDescending(q => Common(open, chains[q.Peek()])).ThenBy(q => Common(next, chains[q.Peek()])).First();
                int row = pick.Dequeue();
                order.Add(row);
                open = chains[row];
            }
            i = end;
        }
        return order;
    }

}

internal sealed record IrComposedTableControls(IrControlLayout? Rows,
    Dictionary<string, (int Row, IrControlLayout Layout)> Cells, List<int>? Order, Dictionary<int, string> TableAnchors);
