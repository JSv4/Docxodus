using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Ir.Diff;

/// <summary>
/// Block-level bookmark and comment-range markers (issue #1149). A <c>w:bookmarkStart</c>/<c>w:bookmarkEnd</c>
/// or <c>w:commentRangeStart</c>/<c>w:commentRangeEnd</c> can sit directly in a block stream (the body, a table
/// cell, a block wrapper's content, a note, a header or footer) as a sibling of <c>w:p</c>; Word writes a bookmark
/// that spans whole paragraphs that way. The IR leaves these markers out of its block list (rule N3), so the edit
/// script never mentions them. The renderer puts them back from the source documents, around the blocks they
/// bracket: each block stream's ops are rendered through <see cref="RenderBlockOpsWordShaped"/>, which places the
/// markers this plan assigns to it.
/// </summary>
internal static partial class IrMarkupRenderer
{
    /// <summary>A bookmark or comment-range marker: the children a block stream holds that are neither a block nor
    /// the stream's own properties, and that the IR drops (rule N3).</summary>
    internal static bool IsBlockLevelMarker(XElement e) =>
        e.Name == W.bookmarkStart || e.Name == W.bookmarkEnd ||
        e.Name == W.commentRangeStart || e.Name == W.commentRangeEnd;

    /// <summary>A bookmark or comment-range marker in a revision wrapper written directly into a block stream: a
    /// marker the block placement (or a replaced wrapper's copy) gave to one document's view only.</summary>
    private static bool IsBlockLevelRevisionMarker(XElement marker) =>
        marker.Parent is { } wrapper && (wrapper.Name == W.ins || wrapper.Name == W.del) &&
        wrapper.Parent is { } stream && (stream.Name == W.body || stream.Name == W.tbl || stream.Name == W.tr || stream.Name == W.tc ||
            stream.Name == W.sdtContent || stream.Name == W.customXml || stream.Name == W.hdr || stream.Name == W.ftr ||
            stream.Name == W.footnote || stream.Name == W.endnote);

    /// <summary>Where the block-level markers of one block stream go, relative to its ops (see
    /// <see cref="PlanBlockLevelMarkers"/>). Every list holds rendered elements: a bare clone, or a clone in a
    /// <c>w:del</c>/<c>w:ins</c>.</summary>
    private sealed class BlockMarkerPlan
    {
        /// <summary>The run of one-sided ops between two ops that have blocks on both sides is a gap; gap
        /// <c>g</c> comes before the <c>g</c>-th two-sided op, the last one after every op. Its start markers go
        /// before the gap's output, its end markers after it.</summary>
        public readonly Dictionary<int, List<XElement>> GapStart = new();

        public readonly Dictionary<int, List<XElement>> GapEnd = new();

        /// <summary>Markers that travel with one op's output, in front of it: a one-sided marker inside a gap,
        /// or one between two blocks the same op consumes.</summary>
        public readonly Dictionary<IrEditOp, List<XElement>> OpPrefix = new(ReferenceEqualityComparer.Instance);

        public List<XElement> PrefixOf(IrEditOp op) => OpPrefix.TryGetValue(op, out var list) ? list : Empty;

        public static List<XElement> Of(Dictionary<int, List<XElement>> map, int key) =>
            map.TryGetValue(key, out var list) ? list : Empty;

        private static readonly List<XElement> Empty = new();
    }

    /// <summary>One side's markers: those before a block (keyed by the op that consumes it and the block's index
    /// among that op's blocks on this side) and those after the last block.</summary>
    private sealed record SideMarkers(Dictionary<(int Op, int Index), List<XElement>> Before, List<XElement> Trailing);

    /// <summary>A position's marker list after merging the two sides: matched markers render bare, the others
    /// as their side's revision.</summary>
    private sealed record MarkerEntry(XElement? Left, XElement? Right);

    /// <summary>
    /// Place the block-level markers of both sides of one block stream around its <paramref name="ops"/>. Each
    /// side's marker sits before one of its blocks or after the last; that fixes a position in the op sequence:
    /// at the start or end of a gap (a run of one-sided ops) or inside one. A marker both sides have at the same
    /// gap boundary is emitted bare, so it survives accept and reject; any other marker is emitted as its side's
    /// revision (<c>w:del</c> for the original's, <c>w:ins</c> for the revised document's), so accept gives the
    /// revised document's markers and reject the original's. A bookmark or comment whose markers are not all
    /// matched keeps one whole copy per side rather than sharing some markers: a shared start with one end per
    /// side could not be paired. Returns null when neither side's stream has a block-level marker.
    /// </summary>
    private static BlockMarkerPlan? PlanBlockLevelMarkers(IReadOnlyList<IrEditOp> ops, RenderState state)
    {
        if (ComputeMarkerLayout(ops, state.Left, state.RightSource) is not { } layout)
            return null;

        List<XElement> Render(List<MarkerEntry> entries) => RenderMarkerEntries(entries, layout.IsShared, state);

        var plan = new BlockMarkerPlan();
        foreach (var (g, entries) in layout.GapStart)
            if (entries.Count > 0)
                plan.GapStart[g] = Render(entries);
        foreach (var (g, entries) in layout.GapEnd)
            if (entries.Count > 0)
                plan.GapEnd[g] = Render(entries);
        foreach (var group in layout.Interior.GroupBy(x => x.Op))
            plan.OpPrefix[ops[group.Key]] = Render(group.Select(x => x.Entry).ToList());
        return plan;
    }

    /// <summary>Both sides' block-level markers of one block stream, merged at the positions of
    /// <see cref="PlanBlockLevelMarkers"/>: gap starts and ends (by gap index) and op interiors (by op index).
    /// <see cref="TwoSidedOps"/> lists the ops that bound the gaps: gap <c>g</c> ends at <c>TwoSidedOps[g]</c>.</summary>
    private sealed record MarkerLayout(
        Dictionary<int, List<MarkerEntry>> GapStart,
        Dictionary<int, List<MarkerEntry>> GapEnd,
        List<(int Op, MarkerEntry Entry)> Interior,
        List<int> TwoSidedOps,
        HashSet<string> Unmatched)
    {
        /// <summary>A marker both sides have at this position, of a bookmark or comment all of whose block-level
        /// markers match: it is emitted once, bare.</summary>
        public bool IsShared(MarkerEntry entry) =>
            entry.Left is not null && entry.Right is not null && !Unmatched.Contains(MarkerIdentity(entry.Right));
    }

    private static MarkerLayout? ComputeMarkerLayout(IReadOnlyList<IrEditOp> ops, IrDocument leftDoc, IrDocument rightDoc)
    {
        // Most streams have no marker between their blocks: find that out from one block per side before
        // resolving every op.
        bool HasMarkers(Func<IrEditOp, string?> anchor, IrDocument doc) =>
            ops.Select(op => SourceElement(anchor(op), doc)?.Parent).FirstOrDefault(p => p is not null) is { } stream &&
            stream.Elements().Any(IsBlockLevelMarker);
        if (!HasMarkers(op => op.LeftAnchor ?? op.SplitMergeAnchors?.FirstOrDefault(), leftDoc) &&
            !HasMarkers(op => op.RightAnchor ?? op.SplitMergeAnchors?.FirstOrDefault(), rightDoc))
            return null;

        var leftBlocks = new List<XElement>[ops.Count];
        var rightBlocks = new List<XElement>[ops.Count];
        for (int i = 0; i < ops.Count; i++)
            (leftBlocks[i], rightBlocks[i]) = OpSourceBlocks(ops[i], leftDoc, rightDoc);

        return ComputeMarkerLayout(leftBlocks, rightBlocks);
    }

    // Rows and cells use the same gap matching as paragraph blocks; their sources are resolved by the table
    // renderer because row/cell anchors are not part of the document's block index.
    private static MarkerLayout? ComputeMarkerLayout(List<XElement>[] leftBlocks, List<XElement>[] rightBlocks)
    {
        var left = ReadSideMarkers(leftBlocks);
        var right = ReadSideMarkers(rightBlocks);
        if (left is null && right is null)
            return null;
        left ??= new SideMarkers(new(), new());
        right ??= new SideMarkers(new(), new());

        // Gap g holds the one-sided ops after the (g-1)-th two-sided op; each side's ops in it, in order.
        var twoSided = new bool[leftBlocks.Length];
        var gapOf = new int[leftBlocks.Length];
        int gaps = 0;
        var leftGapOps = new List<List<int>> { new() };
        var rightGapOps = new List<List<int>> { new() };
        for (int i = 0; i < leftBlocks.Length; i++)
        {
            twoSided[i] = leftBlocks[i].Count > 0 && rightBlocks[i].Count > 0;
            gapOf[i] = gaps;
            if (twoSided[i])
            {
                gaps++;
                leftGapOps.Add(new());
                rightGapOps.Add(new());
                continue;
            }
            if (leftBlocks[i].Count > 0)
                leftGapOps[gaps].Add(i);
            if (rightBlocks[i].Count > 0)
                rightGapOps[gaps].Add(i);
        }

        var interior = new List<(int Op, MarkerEntry Entry)>();

        // Sort each side's markers into gap-start, gap-end and interior buckets.
        (Dictionary<int, List<XElement>> Start, Dictionary<int, List<XElement>> End) Bucket(
            SideMarkers side, List<List<int>> gapOps, bool isLeft)
        {
            var start = new Dictionary<int, List<XElement>>();
            var end = new Dictionary<int, List<XElement>>();
            void Add(Dictionary<int, List<XElement>> map, int g, List<XElement> markers)
            {
                if (!map.TryGetValue(g, out var list))
                    map[g] = list = new List<XElement>();
                list.AddRange(markers);
            }
            foreach (var ((op, index), markers) in side.Before.OrderBy(kv => kv.Key.Op).ThenBy(kv => kv.Key.Index))
            {
                if (index == 0 && twoSided[op])
                    Add(end, gapOf[op], markers);
                else if (index == 0 && gapOps[gapOf[op]].IndexOf(op) == 0)
                    Add(start, gapOf[op], markers);
                else
                    interior.AddRange(markers.Select(m => (op, isLeft ? new MarkerEntry(m, null) : new MarkerEntry(null, m))));
            }
            Add(end, gaps, side.Trailing);
            return (start, end);
        }

        var (leftStart, leftEnd) = Bucket(left, leftGapOps, isLeft: true);
        var (rightStart, rightEnd) = Bucket(right, rightGapOps, isLeft: false);

        // Merge the two sides at each gap boundary. A side with no op in the gap has one boundary there, which
        // the other side's start and end both face: split its markers between them where most of them match.
        var gapStart = new Dictionary<int, List<MarkerEntry>>();
        var gapEnd = new Dictionary<int, List<MarkerEntry>>();
        for (int g = 0; g <= gaps; g++)
        {
            var ls = BlockMarkerPlan.Of(leftStart, g);
            var le = BlockMarkerPlan.Of(leftEnd, g);
            var rs = BlockMarkerPlan.Of(rightStart, g);
            var re = BlockMarkerPlan.Of(rightEnd, g);
            bool leftEmpty = leftGapOps[g].Count == 0, rightEmpty = rightGapOps[g].Count == 0;
            if (leftEmpty && !rightEmpty)
                (gapStart[g], gapEnd[g]) = MergeAcrossSplit(le, rs, re, leftIsSplit: true);
            else if (rightEmpty && !leftEmpty)
                (gapStart[g], gapEnd[g]) = MergeAcrossSplit(re, ls, le, leftIsSplit: false);
            else
            {
                gapStart[g] = MergeMarkers(ls, rs);
                gapEnd[g] = MergeMarkers(le, re);
            }
        }

        // A bookmark or comment with any one-sided marker keeps a whole copy per side: split its matched ones.
        var unmatched = new HashSet<string>(StringComparer.Ordinal);
        foreach (var entry in gapStart.Values.Concat(gapEnd.Values).SelectMany(l => l)
                     .Concat(interior.Select(x => x.Entry)))
            if (entry.Left is null || entry.Right is null)
                unmatched.Add(MarkerIdentity(entry.Left ?? entry.Right!));

        var twoSidedOps = Enumerable.Range(0, leftBlocks.Length).Where(i => twoSided[i]).ToList();
        return new MarkerLayout(gapStart, gapEnd, interior, twoSidedOps, unmatched);
    }

    /// <summary>
    /// The consolidate counterpart of <see cref="PlanBlockLevelMarkers"/> for one base-anchored block stream of
    /// composite <paramref name="ops"/>: the markers to emit before each op (index <c>ops.Count</c>: after the
    /// last). Each base marker stays where the base has it, bare, or as the deletion of the first reviewer whose
    /// own comparison with the base (<paramref name="reviewerOps"/>, that reviewer's two-way ops over the same
    /// stream) does not keep it. Each reviewer's own marker goes in as that reviewer's insertion, at the same
    /// position relative to the base blocks, or before the reviewer's inserted block it precedes. With no
    /// reviewer ops (a stream the merger composed without them) the base markers are kept as they are. Returns
    /// null when there is nothing to place.
    /// </summary>
    internal static Dictionary<int, List<XElement>>? PlanCompositeBlockLevelMarkers(
        IReadOnlyList<IrCompositeOp> ops, IrDocument baseIr, IReadOnlyList<IrDocument> reviewerIrs,
        IReadOnlyList<IReadOnlyList<IrEditOp>>? reviewerOps, IReadOnlyList<string> authors, RenderState state)
    {
        // Nothing to place when the base has no marker between these blocks and no reviewer has one between
        // theirs: decide that before resolving every op.
        bool baseHasMarkers =
            ops.Select(o => SourceElement(o.Op.LeftAnchor, baseIr)?.Parent).FirstOrDefault(p => p is not null) is { } stream &&
            stream.Elements().Any(IsBlockLevelMarker);
        var layouts = new MarkerLayout?[reviewerOps?.Count ?? 0];
        for (int r = 0; r < layouts.Length && r < reviewerIrs.Count; r++)
            layouts[r] = ComputeMarkerLayout(reviewerOps![r], baseIr, reviewerIrs[r]);
        if (!baseHasMarkers && layouts.All(l => l is null))
            return null;
        var baseBlocks = ops.Select(o => SourceElement(o.Op.LeftAnchor, baseIr) is { } el
            ? new List<XElement> { el } : new List<XElement>()).ToArray();
        var baseSide = ReadSideMarkers(baseBlocks);

        var byBaseAnchor = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int i = 0; i < ops.Count; i++)
            if (ops[i].Op.LeftAnchor is { } a)
                byBaseAnchor.TryAdd(a, i);
        int? BaseIndex(string? anchor) => anchor is not null && byBaseAnchor.TryGetValue(anchor, out var i) ? i : null;

        return PlanCompositeMarkers(ops.Count, baseSide, layouts,
            (r, j) => BaseIndex(reviewerOps![r][j].LeftAnchor),
            (r, j) => reviewerOps![r][j].RightAnchor is { } right
                ? Enumerable.Range(0, ops.Count).Cast<int?>().FirstOrDefault(i =>
                    ops[i!.Value].SourceReviewer == r && ops[i.Value].Op.RightAnchor == right)
                : null, authors, state);
    }

    private static Dictionary<int, List<XElement>>? PlanCompositeMarkers(int count, SideMarkers? baseSide,
        IReadOnlyList<MarkerLayout?> layouts, Func<int, int, int?> baseIndex, Func<int, int, int?> rightIndex,
        IReadOnlyList<string> authors, RenderState state)
    {
        var placement = new CompositeMarkerPlacement(state);
        for (int r = 0; r < layouts.Count; r++)
        {
            if (layouts[r] is not { } layout)
                continue;
            var two = layout.TwoSidedOps;
            int reviewer = r;
            void Take(MarkerEntry entry, int position, bool beforeBase)
            {
                if (!layout.IsShared(entry))
                    placement.Take(entry, reviewer, position, beforeBase);
            }
            foreach (var (g, entries) in layout.GapStart)
            {
                int position = g == 0 ? 0 : (baseIndex(r, two[g - 1]) ?? count - 1) + 1;
                foreach (var entry in entries)
                    Take(entry, position, beforeBase: true);
            }
            foreach (var (g, entries) in layout.GapEnd)
            {
                int position = g < two.Count ? baseIndex(r, two[g]) ?? count : count;
                foreach (var entry in entries)
                    Take(entry, position, beforeBase: false);
            }
            foreach (var (j, entry) in layout.Interior)
            {
                int position = rightIndex(r, j) ?? baseIndex(r, j) ?? count;
                Take(entry, position, beforeBase: false);
            }
        }

        if (baseSide is null && placement.IsEmpty)
            return null;

        var plan = new Dictionary<int, List<XElement>>();
        for (int i = 0; i <= count; i++)
        {
            var baseMarkers = baseSide is null ? new List<XElement>()
                : i == count ? baseSide.Trailing : baseSide.Before.GetValueOrDefault((i, 0)) ?? new List<XElement>();
            var output = placement.Render(i, baseMarkers, authors, state);
            if (output.Count > 0)
                plan[i] = output;
        }
        return plan;
    }

    /// <summary>Put a table or row's markers around its rendered rows/cells, using their actual pairings rather
    /// than their output ordinals. An inserted/deleted row or cell can shift every subsequent boundary.</summary>
    private static void ReconcileTableStreamMarkers(XElement emitted,
        IReadOnlyList<(XElement? Left, XElement? Right)> pairs, RenderState state)
    {
        var layout = ComputeMarkerLayout(pairs.Select(p => MarkerBlock(p.Left)).ToArray(),
            pairs.Select(p => MarkerBlock(p.Right)).ToArray());
        if (layout is null)
            return;
        var entries = new Dictionary<int, List<MarkerEntry>>();
        void Add(int index, IEnumerable<MarkerEntry> values)
        {
            if (!entries.TryGetValue(index, out var list))
                entries[index] = list = new();
            list.AddRange(values);
        }
        foreach (var (gap, values) in layout.GapStart)
            Add(gap == 0 ? 0 : layout.TwoSidedOps[gap - 1] + 1, values);
        foreach (var (gap, values) in layout.GapEnd)
            Add(gap < layout.TwoSidedOps.Count ? layout.TwoSidedOps[gap] : pairs.Count, values);
        foreach (var (index, entry) in layout.Interior)
            Add(index, new[] { entry });
        PlaceStreamMarkers(emitted, entries.ToDictionary(p => p.Key,
            p => RenderMarkerEntries(p.Value, layout.IsShared, state)));
    }

    private static List<XElement> MarkerBlock(XElement? element) => element is null ? new() : new() { element };

    private static void PlaceStreamMarkers(XElement stream, IReadOnlyDictionary<int, List<XElement>> markers)
    {
        var blocks = stream.Elements().Where(IsMarkerStreamContent).ToList();
        stream.Elements().Where(IsBlockLevelMarker).Remove();
        foreach (var (index, values) in markers.OrderBy(p => p.Key))
        {
            if (index < blocks.Count)
                blocks[index].AddBeforeSelf(values);
            else
                stream.Add(values);
        }
    }

    internal static bool HasTableBoundaryMarkers(XElement? table) => table is not null &&
        (table.Elements().Any(IsBlockLevelMarker) || table.Elements(W.tr).Any(r => r.Elements().Any(IsBlockLevelMarker)));

    private sealed record MarkerStreamReview(int Reviewer, string Author,
        IReadOnlyList<(XElement? Left, XElement? Right)> Pairs);

    private static void ReconcileCompositeTableStreamMarkers(XElement emitted, IReadOnlyList<XElement?> baseElements,
        IReadOnlyList<(int Reviewer, XElement? Element)> insertedElements, IReadOnlyList<MarkerStreamReview> reviewers,
        int reviewerCount, RenderState state)
    {
        var baseSide = ReadSideMarkers(baseElements.Select(MarkerBlock).ToArray());
        var layouts = new MarkerLayout?[reviewerCount];
        var authors = new string[reviewerCount];
        foreach (var reviewer in reviewers)
        {
            layouts[reviewer.Reviewer] = ComputeMarkerLayout(reviewer.Pairs.Select(p => MarkerBlock(p.Left)).ToArray(),
                reviewer.Pairs.Select(p => MarkerBlock(p.Right)).ToArray());
            authors[reviewer.Reviewer] = reviewer.Author;
        }
        if (baseSide is null && layouts.All(l => l is null))
            return;
        var byBase = baseElements.Select((element, index) => (element, index)).Where(p => p.element is not null)
            .ToDictionary(p => p.element!, p => p.index);
        var byRight = insertedElements.Select((pair, index) => (pair, index)).Where(p => p.pair.Element is not null)
            .ToDictionary(p => p.pair, p => p.index);
        XElement? Source(int r, int j, bool right)
        {
            var pair = reviewers.First(v => v.Reviewer == r).Pairs[j];
            return right ? pair.Right : pair.Left;
        }
        var plan = PlanCompositeMarkers(baseElements.Count, baseSide, layouts,
            (r, j) => Source(r, j, false) is { } left && byBase.TryGetValue(left, out var i) ? i : null,
            (r, j) => Source(r, j, true) is { } right && byRight.TryGetValue((r, right), out var i) ? i : null,
            authors, state);
        if (plan is not null)
            PlaceStreamMarkers(emitted, plan);
    }

    private static void ReconcileComposedTableMarkers(XElement emitted, IrCompositeOp op,
        Dictionary<string, XElement> baseRows, Dictionary<string, XElement> baseCells,
        IReadOnlyList<IrDocument> reviewerIrs, RenderState state)
    {
        if (op.AuthoredRows is not { } rows || op.ReviewerTables is not { } tables)
            return;
        XElement? RightRow(int reviewer, string? anchor) => anchor is null ? null : FindRowSource(reviewerIrs[reviewer], anchor);
        var reviews = tables.Select(t => new MarkerStreamReview(t.Reviewer, t.Author,
            t.Diff.RowOps.Select(r => (
                r.Kind == IrRowOpKind.MovedRow && r.IsMoveSource != true ? null : baseRows.GetValueOrDefault(r.LeftRowAnchor ?? ""),
                r.Kind == IrRowOpKind.MovedRow && r.IsMoveSource == true ? null : RightRow(t.Reviewer, r.RightRowAnchor)))
                .ToList())).ToList();
        ReconcileCompositeTableStreamMarkers(emitted,
            rows.Select(r => baseRows.GetValueOrDefault(r.BaseRowAnchor ?? "")).ToList(),
            rows.Select(r => (r.SourceReviewer, r.Kind == IrRowOpKind.InsertRow ? RightRow(r.SourceReviewer, r.RightRowAnchor) : null)).ToList(),
            reviews, reviewerIrs.Count, state);

        var outputRows = emitted.Elements(W.tr).ToList();
        for (int i = 0; i < rows.Count && i < outputRows.Count; i++)
        {
            var row = rows[i];
            if (row.Kind != IrRowOpKind.EqualRow && row.Kind != IrRowOpKind.ModifyRow ||
                row.BaseRowAnchor is not { } anchor || !baseRows.TryGetValue(anchor, out var source))
                continue;
            var baseRowCells = source.Elements(W.tc).ToList();
            var cellReviews = new List<MarkerStreamReview>();
            foreach (var table in tables)
            {
                var pair = table.Diff.RowOps.FirstOrDefault(r => r.LeftRowAnchor == anchor);
                if (pair is null || RightRow(table.Reviewer, pair.RightRowAnchor) is not { } right)
                    continue;
                var rightCells = right.Elements(W.tc).ToList();
                int l = 0, r = 0;
                var pairs = pair.CellOps is { } cells
                    ? cells.Select(c => (c.LeftCellAnchor is null ? null : baseRowCells[l++],
                        c.RightCellAnchor is null ? null : rightCells[r++])).ToList()
                    : baseRowCells.Select((c, j) => ((XElement?)c, j < rightCells.Count ? rightCells[j] : null)).ToList();
                cellReviews.Add(new MarkerStreamReview(table.Reviewer, table.Author, pairs));
            }
            var outputBase = row.ComposedCells is { } composed
                ? composed.Select(c => baseCells.GetValueOrDefault(c.BaseCellAnchor ?? "")).ToList()
                : baseRowCells.Cast<XElement?>().ToList();
            var outputInserted = row.ComposedCells is { } cc
                ? cc.Select(c => (c.ShellSourceReviewer, c.Kind == IrAuthoredCellKind.InsertCell && c.ShellRightCellAnchor is { } right
                    ? FindCellSource(reviewerIrs[c.ShellSourceReviewer], right) : null)).ToList()
                : baseRowCells.Select(_ => (-1, (XElement?)null)).ToList();
            ReconcileCompositeTableStreamMarkers(outputRows[i], outputBase, outputInserted, cellReviews, reviewerIrs.Count, state);
        }
    }

    /// <summary>
    /// The N-way rule for one base block stream's markers (issue #1149, #1175), shared by the composite plan and
    /// <see cref="ReconcileCompositeEqualContainerMarkers"/>. Each reviewer's comparison with the base yields
    /// merged marker entries per base block boundary (position); a shared entry is that reviewer keeping the base
    /// marker. A base marker stays bare unless a reviewer does not keep it: it is then the deletion of the first
    /// such reviewer. A reviewer's own marker goes in as that reviewer's insertion at the boundary, in front of
    /// the base markers there (the start of a gap) or after them; a marker another reviewer already brought is
    /// not repeated.
    /// </summary>
    private sealed class CompositeMarkerPlacement
    {
        private readonly Dictionary<XElement, int> _deletedBy = new(ReferenceEqualityComparer.Instance);
        private readonly Dictionary<(XElement? Stream, int Position), List<(int Reviewer, XElement Marker)>> _beforeBase = new();
        private readonly Dictionary<(XElement? Stream, int Position), List<(int Reviewer, XElement Marker)>> _afterBase = new();
        private readonly HashSet<string> _added;

        /// <summary>Markers a reviewer brings are recorded package-wide (<see cref="RenderState.CompositeAddedMarkers"/>):
        /// a bookmark two reviewers add under one name, in different block streams, is inserted once.</summary>
        public CompositeMarkerPlacement(RenderState state) =>
            _added = state.CompositeAddedMarkers ??= new HashSet<string>(StringComparer.Ordinal);

        public bool IsEmpty => _deletedBy.Count == 0 && _beforeBase.Count == 0 && _afterBase.Count == 0;

        /// <summary>Record one reviewer's entry that is not shared with the base, at <paramref name="position"/>
        /// of <paramref name="stream"/> (null for a plan over one stream's ops).</summary>
        public void Take(MarkerEntry entry, int reviewer, int position, bool beforeBase, XElement? stream = null)
        {
            if (entry.Left is not null)
                _deletedBy.TryAdd(entry.Left, reviewer);
            if (entry.Right is null || !_added.Add(MarkerKey(entry.Right)))
                return; // another reviewer already brought the same marker
            var map = beforeBase ? _beforeBase : _afterBase;
            if (!map.TryGetValue((stream, position), out var list))
                map[(stream, position)] = list = new List<(int, XElement)>();
            list.Add((reviewer, entry.Right));
        }

        /// <summary>The markers to emit at <paramref name="position"/>, whose base markers are
        /// <paramref name="baseMarkers"/>: each as a bare clone or as a reviewer's revision.</summary>
        public List<XElement> Render(int position, IReadOnlyList<XElement> baseMarkers, IReadOnlyList<string> authors,
            RenderState state, XElement? stream = null)
        {
            var savedAuthor = state.AuthorOverride;
            XElement Revision(XName name, int reviewer, XElement marker)
            {
                state.AuthorOverride = reviewer >= 0 && reviewer < authors.Count ? authors[reviewer] : savedAuthor;
                return new XElement(name, state.RevisionAttributes(), StripUnids(new XElement(marker)));
            }
            var output = new List<XElement>();
            foreach (var (reviewer, marker) in _beforeBase.GetValueOrDefault((stream, position)) ?? new())
                output.Add(Revision(W.ins, reviewer, marker));
            foreach (var marker in baseMarkers)
                output.Add(_deletedBy.TryGetValue(marker, out var reviewer)
                    ? Revision(W.del, reviewer, marker)
                    : StripUnids(new XElement(marker)));
            foreach (var (reviewer, marker) in _afterBase.GetValueOrDefault((stream, position)) ?? new())
                output.Add(Revision(W.ins, reviewer, marker));
            state.AuthorOverride = savedAuthor;
            return output;
        }
    }

    /// <summary>
    /// The consolidate counterpart of <see cref="ReconcileEqualContainerMarkers"/>: put each reviewer's
    /// block-level markers into a container emitted from the base because every reviewer left it equal (an
    /// unchanged table, cell, block wrapper or note; issue #1175). <paramref name="counterparts"/> are the
    /// reviewers' copies of <paramref name="baseSource"/>, which line up with it block for block. Each block
    /// boundary of every stream inside is merged per reviewer as the two-way reconcile merges it, then placed by
    /// the N-way rule (<see cref="CompositeMarkerPlacement"/>): a base marker a reviewer drops becomes that
    /// reviewer's deletion, a reviewer's own marker that reviewer's insertion.
    /// </summary>
    internal static void ReconcileCompositeEqualContainerMarkers(XElement emitted, XElement baseSource,
        IReadOnlyList<(int Reviewer, XElement Counterpart)> counterparts, IReadOnlyList<string> authors, RenderState state)
    {
        var placement = new CompositeMarkerPlacement(state);
        var streams = new Dictionary<XElement, XElement>(ReferenceEqualityComparer.Instance); // emitted -> base
        foreach (var (reviewer, counterpart) in counterparts)
        {
            if (ReferenceEquals(counterpart, baseSource) ||
                !HasMarkerBetweenBlocks(baseSource) && !HasMarkerBetweenBlocks(counterpart))
                continue;
            var entries = new List<(XElement Stream, XElement LeftStream, int Boundary, List<MarkerEntry> Entries)>();
            CollectEqualStreamMarkers(emitted, baseSource, counterpart, entries);
            var unmatched = entries.SelectMany(e => e.Entries).Where(e => e.Left is null || e.Right is null)
                .Select(e => MarkerIdentity(e.Left ?? e.Right!)).ToHashSet(StringComparer.Ordinal);
            if (unmatched.Count == 0)
                continue; // this reviewer kept every base marker and added none
            foreach (var (stream, leftStream, boundary, list) in entries)
            {
                streams.TryAdd(stream, leftStream);
                foreach (var entry in list)
                    if (entry.Left is null || entry.Right is null || unmatched.Contains(MarkerIdentity(entry.Right)))
                        placement.Take(entry, reviewer, boundary, beforeBase: false, stream);
            }
        }
        if (placement.IsEmpty)
            return;
        foreach (var (stream, leftStream) in streams)
        {
            var blocks = stream.Elements().Where(IsMarkerStreamContent).ToList();
            var (_, baseBefore) = SplitStream(leftStream);
            if (baseBefore.Count != blocks.Count + 1)
                continue;
            stream.Elements().Where(IsBlockLevelMarker).Remove();
            for (int k = 0; k <= blocks.Count; k++)
            {
                var rendered = placement.Render(k, baseBefore[k], authors, state, stream);
                if (rendered.Count == 0)
                    continue;
                if (k < blocks.Count)
                    blocks[k].AddBeforeSelf(rendered);
                else if (blocks.Count > 0)
                    blocks[^1].AddAfterSelf(rendered);
                else
                    stream.Add(rendered);
            }
        }
    }

    /// <summary>
    /// Put both sides' block-level markers into a block container the renderer emitted as one side's copy
    /// because the IR found the two equal: an unchanged table, row, cell or block wrapper. The IR leaves markers
    /// out (rule N3), so a container whose only difference is a bookmark or comment range between its blocks
    /// pairs equal and is cloned from the revised document, which loses the original's markers and leaves the
    /// revised document's bare. Equal containers have the same blocks in the same order, so the k-th block of
    /// each side is the same block: each boundary's markers are merged as <see cref="PlanBlockLevelMarkers"/>
    /// merges a gap boundary, replacing the copy's own, and nested tables and wrappers are visited the same way.
    /// Containers whose shapes do not line up (a copy carrying preserved input revisions) are left as emitted.
    /// </summary>
    internal static bool ReconcileEqualContainerMarkers(XElement emitted, XElement? left, XElement? right, RenderState state)
    {
        if (left is null || right is null || ReferenceEquals(left, right) ||
            !HasMarkerBetweenBlocks(left) && !HasMarkerBetweenBlocks(right))
            return false;
        var entries = new List<(XElement Stream, XElement LeftStream, int Boundary, List<MarkerEntry> Entries)>();
        CollectEqualStreamMarkers(emitted, left, right, entries);
        if (entries.Count == 0)
            return false;

        var unmatched = new HashSet<string>(StringComparer.Ordinal);
        foreach (var entry in entries.SelectMany(e => e.Entries))
            if (entry.Left is null || entry.Right is null)
                unmatched.Add(MarkerIdentity(entry.Left ?? entry.Right!));
        if (unmatched.Count == 0)
            return false; // every marker matched: the copy already holds both sides' markers
        foreach (var group in entries.GroupBy(e => e.Stream))
        {
            var stream = group.Key;
            var blocks = stream.Elements().Where(IsMarkerStreamContent).ToList();
            stream.Elements().Where(IsBlockLevelMarker).Remove();
            foreach (var (_, _, boundary, list) in group)
            {
                var rendered = RenderMarkerEntries(list, entry => entry.Left is not null && entry.Right is not null &&
                    !unmatched.Contains(MarkerIdentity(entry.Right)), state);
                if (boundary < blocks.Count)
                    blocks[boundary].AddBeforeSelf(rendered);
                else if (blocks.Count > 0)
                    blocks[^1].AddAfterSelf(rendered);
                else
                    stream.Add(rendered);
            }
        }
        return true;
    }

    /// <summary>
    /// Whether two block containers the IR found equal differ in their block-level markers: the bookmarks and
    /// comment ranges each block stream inside holds between its blocks, at each block boundary. A container is a
    /// story root (header, footer, note), a cell or custom-XML wrapper (each its own stream), a table (its cells)
    /// or a content control (its content); nested tables and wrappers are compared the same way. Equal containers
    /// have the same blocks, so the streams line up; markers are compared as <see cref="MergeMarkers"/> matches
    /// them. A missing container differs from nothing (issues #1174, #1175).
    /// </summary>
    internal static bool BlockLevelMarkersDiffer(XElement? left, XElement? right,
        IrDocument? leftDocument = null, IrDocument? rightDocument = null)
    {
        if (left is null || right is null || !HasStreamMarker(left) && !HasStreamMarker(right))
            return false;
        // The comment identities are read only now that a marker between blocks is known to exist.
        var (leftComments, rightComments) = (leftDocument is null ? null : CommentIdentities(leftDocument),
            rightDocument is null ? null : CommentIdentities(rightDocument));
        return !ContainerSignature(left, leftComments).SequenceEqual(ContainerSignature(right, rightComments), StringComparer.Ordinal);

        // Most containers hold no marker in a block stream: tell that without building signatures.
        static bool HasStreamMarker(XElement container) =>
            container.DescendantsAndSelf().Any(e => IsBlockLevelMarker(e) && e.Parent is { } p && p.Name != W.p &&
                (p == container || p.Name == W.tbl || p.Name == W.tr || p.Name == W.tc || p.Name == W.sdtContent || p.Name == W.customXml));

        static IEnumerable<string> ContainerSignature(XElement container, IReadOnlyDictionary<string, string>? comments)
        {
            if (container.Name == W.sdt)
                return BlockWrapperContent(container) is { } content ? StreamSignature(content, comments) : Enumerable.Empty<string>();
            return StreamSignature(container, comments);
        }

        // A comment is matched by what it says, not its id: the revised document may number the same comment
        // differently (another comment added before it), which is no change here.
        static string Key(XElement marker, IReadOnlyDictionary<string, string>? comments) =>
            (marker.Name == W.commentRangeStart || marker.Name == W.commentRangeEnd) &&
            comments is not null && (string?)marker.Attribute(W.id) is { } id && comments.TryGetValue(id, out var identity)
                ? marker.Name.LocalName + "\u0000c\u0000" + identity
                : MarkerKey(marker);

        // A stream's markers and block boundaries, depth first; a nested container is bracketed so its markers
        // never line up with the outer stream's.
        static IEnumerable<string> StreamSignature(XElement stream, IReadOnlyDictionary<string, string>? comments)
        {
            foreach (var child in stream.Elements())
            {
                if (IsBlockLevelMarker(child))
                {
                    yield return Key(child, comments);
                    continue;
                }
                if (!IsMarkerStreamContent(child))
                    continue;
                yield return "|";
                if (child.Name == W.p)
                    continue;
                yield return "(";
                foreach (var token in ContainerSignature(child, comments))
                    yield return token;
                yield return ")";
            }
        }
    }

    /// <summary>Each comment of <paramref name="document"/> by id, as what identifies it across documents: its author,
    /// date and text. Empty when the IR kept no source for the comments part. Built once per document.</summary>
    internal static IReadOnlyDictionary<string, string> CommentIdentities(IrDocument document) =>
        CommentIdentityMaps.GetValue(document, BuildCommentIdentities);

    private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<IrDocument, IReadOnlyDictionary<string, string>>
        CommentIdentityMaps = new();

    private static IReadOnlyDictionary<string, string> BuildCommentIdentities(IrDocument document)
    {
        var identities = new Dictionary<string, string>(StringComparer.Ordinal);
        if (document.Comments.PartUri is not { } uri || !document.Sources.TryGetValue(uri, out var source) || source.Root is null)
            return identities;
        foreach (var comment in source.Root.Elements(W.comment))
            if ((string?)comment.Attribute(W.id) is { } id)
                identities.TryAdd(id, (string?)comment.Attribute(W.author) + "\u0000" + (string?)comment.Attribute(W.date) + "\u0000" +
                    string.Concat(comment.Descendants(W.t).Select(t => t.Value)));
        return identities;
    }

    /// <summary>Whether <paramref name="container"/> (itself included) holds a marker directly in a block stream
    /// below it: a cell, a block wrapper's content or a note.</summary>
    private static bool HasMarkerBetweenBlocks(XElement container) =>
        container.DescendantsAndSelf().Any(e => IsBlockLevelMarker(e) && e.Parent is { } p &&
            (p.Name == W.tbl || p.Name == W.tr || p.Name == W.tc || p.Name == W.sdtContent || p.Name == W.customXml || p.Name == W.footnote || p.Name == W.endnote));

    /// <summary>
    /// The output keeps the original's copy of every note the comparison did not change, so a note whose only
    /// difference is a bookmark or comment range between its blocks (invisible to the IR, rule N3) kept the
    /// original's markers on accept too. A note the comparison left alone has the same id in both documents (a
    /// renumbered one gets a diff), so each such note is reconciled against the revised document's note of that
    /// id, as an equal container is (<see cref="ReconcileEqualContainerMarkers"/>).
    /// </summary>
    private static void ReconcileUnchangedNoteMarkers(
        MainDocumentPart main, MainDocumentPart? rightMain, IReadOnlyList<IrNoteDiff>? noteOps, RenderState state)
    {
        foreach (var (part, rightPart, name, kind) in new (OpenXmlPart?, OpenXmlPart?, XName, IrNoteKind)[]
                 {
                     (main.FootnotesPart, rightMain?.FootnotesPart, W.footnote, IrNoteKind.Footnote),
                     (main.EndnotesPart, rightMain?.EndnotesPart, W.endnote, IrNoteKind.Endnote),
                 })
        {
            if (part?.GetXDocument().Root is not { } root || rightPart?.GetXDocument().Root is not { } rightRoot ||
                !HasMarkerBetweenBlocks(root) && !HasMarkerBetweenBlocks(rightRoot))
                continue;
            var touched = (noteOps ?? Array.Empty<IrNoteDiff>()).Where(d => d.Kind == kind)
                .SelectMany(d => new[] { d.NoteId, d.LeftNoteId }).OfType<string>().ToHashSet(StringComparer.Ordinal);
            var rightById = rightRoot.Elements(name).Where(n => n.Attribute(W.id) is not null)
                .GroupBy(n => (string)n.Attribute(W.id)!).ToDictionary(g => g.Key, g => g.First(), StringComparer.Ordinal);
            bool changed = false;
            foreach (var note in root.Elements(name).ToList())
            {
                if ((string?)note.Attribute(W.id) is not { } id || touched.Contains(id) ||
                    !rightById.TryGetValue(id, out var rightNote) ||
                    !HasMarkerBetweenBlocks(note) && !HasMarkerBetweenBlocks(rightNote))
                    continue;
                // The output note is the original's copy; reconcile it against a snapshot of itself.
                changed |= ReconcileEqualContainerMarkers(note, new XElement(note), rightNote, state);
            }
            if (changed)
                part.PutXDocument();
        }
    }

    /// <summary>The merged markers at each block boundary of every block stream inside an equal container pair
    /// (see <see cref="ReconcileEqualContainerMarkers"/>), for the boundaries where the sides differ.</summary>
    private static void CollectEqualStreamMarkers(XElement emitted, XElement left, XElement right,
        List<(XElement Stream, XElement LeftStream, int Boundary, List<MarkerEntry> Entries)> sink)
    {
        if (emitted.Name != left.Name || emitted.Name != right.Name)
            return;
        var stream = emitted.Name == W.sdt ? emitted.Element(W.sdtContent) : emitted;
        var leftStream = left.Name == W.sdt ? left.Element(W.sdtContent) : left;
        var rightStream = right.Name == W.sdt ? right.Element(W.sdtContent) : right;
        if (stream is null || leftStream is null || rightStream is null ||
            !(emitted.Name == W.tbl || emitted.Name == W.tr || emitted.Name == W.tc || emitted.Name == W.sdt || emitted.Name == W.customXml ||
              emitted.Name == W.footnote || emitted.Name == W.endnote))
            return;
        // A copy holding a revision wrapper between its blocks (preserved input revisions) is not a plain clone.
        if (stream.Elements().Any(c => c.Name == W.ins || c.Name == W.del))
            return;

        if (!leftStream.Elements().Any(IsBlockLevelMarker) && !rightStream.Elements().Any(IsBlockLevelMarker))
        {
            // No marker at this level: only nested containers can hold one.
            var (en, ln, rn) = (stream.Elements().Where(IsMarkerStreamContent).ToList(),
                leftStream.Elements().Where(IsMarkerStreamContent).ToList(), rightStream.Elements().Where(IsMarkerStreamContent).ToList());
            if (en.Count == ln.Count && en.Count == rn.Count)
                for (int k = 0; k < en.Count; k++)
                    if (en[k].Name != W.p)
                        CollectEqualStreamMarkers(en[k], ln[k], rn[k], sink);
            return;
        }

        var (eb, _) = SplitStream(stream);
        var (lb, lm) = SplitStream(leftStream);
        var (rb, rm) = SplitStream(rightStream);
        if (eb.Count != lb.Count || eb.Count != rb.Count)
            return;
        for (int k = 0; k <= eb.Count; k++)
            if (lm[k].Count > 0 || rm[k].Count > 0)
                sink.Add((stream, leftStream, k, MergeMarkers(lm[k], rm[k])));
        for (int k = 0; k < eb.Count; k++)
            if (eb[k].Name != W.p)
                CollectEqualStreamMarkers(eb[k], lb[k], rb[k], sink);
    }

    private static bool IsMarkerStreamContent(XElement element) =>
        IsBlockContent(element) || element.Name == W.tr || element.Name == W.tc;

    /// <summary>A block stream's blocks and, for each block boundary (before each block, then after the last), the
    /// block-level markers there.</summary>
    private static (List<XElement> Blocks, List<List<XElement>> Before) SplitStream(XElement stream)
    {
        var blocks = new List<XElement>();
        var before = new List<List<XElement>> { new() };
        foreach (var c in stream.Elements())
        {
            if (IsBlockLevelMarker(c))
                before[^1].Add(c);
            else if (IsMarkerStreamContent(c))
            {
                blocks.Add(c);
                before.Add(new List<XElement>());
            }
        }
        return (blocks, before);
    }

    /// <summary>Render merged marker entries: a shared one bare (the revised document's copy), any other as its
    /// side's revision.</summary>
    private static List<XElement> RenderMarkerEntries(IEnumerable<MarkerEntry> entries, Func<MarkerEntry, bool> isShared,
        RenderState state)
    {
        var output = new List<XElement>();
        foreach (var entry in entries)
        {
            if (isShared(entry))
            {
                output.Add(StripUnids(new XElement(entry.Right!)));
                continue;
            }
            if (entry.Left is not null)
                output.Add(new XElement(W.del, state.RevisionAttributes(), StripUnids(new XElement(entry.Left))));
            if (entry.Right is not null)
                output.Add(new XElement(W.ins, state.RevisionAttributes(), StripUnids(new XElement(entry.Right))));
        }
        return output;
    }

    /// <summary>The source blocks an op consumes on each side, in document order.</summary>
    private static (List<XElement> Left, List<XElement> Right) OpSourceBlocks(IrEditOp op, RenderState state) =>
        OpSourceBlocks(op, state.Left, state.RightSource);

    private static (List<XElement> Left, List<XElement> Right) OpSourceBlocks(IrEditOp op, IrDocument leftDoc, IrDocument rightDoc)
    {
        var left = new List<XElement>();
        var right = new List<XElement>();
        void Add(List<XElement> list, string? anchor, IrDocument doc)
        {
            if (SourceElement(anchor, doc) is { } el && !list.Contains(el))
                list.Add(el);
        }
        switch (op.Kind)
        {
            case IrEditOpKind.SplitBlock:
                Add(left, op.LeftAnchor, leftDoc);
                foreach (var a in (IEnumerable<string>?)op.SplitMergeAnchors ?? Array.Empty<string>())
                    Add(right, a, rightDoc);
                break;
            case IrEditOpKind.MergeBlock:
                foreach (var a in (IEnumerable<string>?)op.SplitMergeAnchors ?? Array.Empty<string>())
                    Add(left, a, leftDoc);
                Add(right, op.RightAnchor, rightDoc);
                break;
            case IrEditOpKind.CrossParagraphRunBlock when op.CrossParagraphCells is { } cells:
                foreach (var cell in cells)
                {
                    Add(left, cell.LeftAnchor, leftDoc);
                    Add(right, cell.RightAnchor, rightDoc);
                }
                break;
            default:
                Add(left, op.LeftAnchor, leftDoc);
                Add(right, op.RightAnchor, rightDoc);
                break;
        }
        return (left, right);
    }

    /// <summary>One side's block-level markers in the stream that holds its <paramref name="blocks"/>: every run
    /// of markers keyed by the consumed block that follows it, and those after the last consumed block. Null
    /// when the side consumes no block or its stream has no marker.</summary>
    private static SideMarkers? ReadSideMarkers(List<XElement>[] blocks)
    {
        var stream = blocks.SelectMany(b => b).Select(b => b.Parent).FirstOrDefault(p => p is not null);
        if (stream is null || !stream.Elements().Any(IsBlockLevelMarker))
            return null;

        var consumed = new Dictionary<XElement, (int Op, int Index)>();
        for (int i = 0; i < blocks.Length; i++)
            for (int k = 0; k < blocks[i].Count; k++)
                if (blocks[i][k].AncestorsAndSelf().FirstOrDefault(a => a.Parent == stream) is { } child)
                    consumed.TryAdd(child, (i, k));

        var before = new Dictionary<(int Op, int Index), List<XElement>>();
        var pending = new List<XElement>();
        foreach (var child in stream.Elements())
        {
            if (IsBlockLevelMarker(child))
            {
                pending.Add(child);
                continue;
            }
            if (pending.Count > 0 && consumed.TryGetValue(child, out var key))
            {
                before[key] = pending;
                pending = new List<XElement>();
            }
        }
        return new SideMarkers(before, pending);
    }

    /// <summary>The markers both sides have at one boundary, in an order both sides keep (a longest common
    /// subsequence of their keys), with each side's others between them in that side's order.</summary>
    private static List<MarkerEntry> MergeMarkers(IReadOnlyList<XElement> left, IReadOnlyList<XElement> right)
    {
        var output = new List<MarkerEntry>();
        if (left.Count == 0 || right.Count == 0)
        {
            output.AddRange(left.Select(l => new MarkerEntry(l, null)));
            output.AddRange(right.Select(r => new MarkerEntry(null, r)));
            return output;
        }
        var lk = left.Select(MarkerKey).ToArray();
        var rk = right.Select(MarkerKey).ToArray();
        var lcs = new int[lk.Length + 1, rk.Length + 1];
        for (int i = lk.Length - 1; i >= 0; i--)
            for (int j = rk.Length - 1; j >= 0; j--)
                lcs[i, j] = lk[i] == rk[j] ? lcs[i + 1, j + 1] + 1 : Math.Max(lcs[i + 1, j], lcs[i, j + 1]);
        int a = 0, b = 0;
        while (a < lk.Length && b < rk.Length)
        {
            if (lk[a] == rk[b])
                output.Add(new MarkerEntry(left[a++], right[b++]));
            else if (lcs[a + 1, b] >= lcs[a, b + 1])
                output.Add(new MarkerEntry(left[a++], null));
            else
                output.Add(new MarkerEntry(null, right[b++]));
        }
        while (a < lk.Length)
            output.Add(new MarkerEntry(left[a++], null));
        while (b < rk.Length)
            output.Add(new MarkerEntry(null, right[b++]));
        return output;
    }

    /// <summary>Merge one side's markers at a boundary that faces both the start and the end of the other side's
    /// run of ops: its first markers go to the start and the rest to the end, split where the most of them match
    /// the other side's.</summary>
    private static (List<MarkerEntry> Start, List<MarkerEntry> End) MergeAcrossSplit(
        List<XElement> split, List<XElement> otherStart, List<XElement> otherEnd, bool leftIsSplit)
    {
        List<MarkerEntry> Merge(IReadOnlyList<XElement> mine, IReadOnlyList<XElement> other) =>
            leftIsSplit ? MergeMarkers(mine, other) : MergeMarkers(other, mine);
        static int Matched(List<MarkerEntry> entries) => entries.Count(e => e.Left is not null && e.Right is not null);

        var best = (Start: Merge(split, otherStart), End: Merge(Array.Empty<XElement>(), otherEnd));
        int bestScore = Matched(best.Start);
        for (int k = split.Count - 1; k >= 0; k--)
        {
            var candidate = (Start: Merge(split.Take(k).ToList(), otherStart), End: Merge(split.Skip(k).ToList(), otherEnd));
            int score = Matched(candidate.Start) + Matched(candidate.End);
            if (score > bestScore)
                (best, bestScore) = (candidate, score);
        }
        return best;
    }

    /// <summary>What makes a marker on one side the same marker as one on the other: its kind and the bookmark's
    /// name (an end's through its start in the same document) or the comment's id.</summary>
    private static string MarkerKey(XElement marker) =>
        marker.Name.LocalName + "\u0000" + MarkerIdentity(marker);

    /// <summary>The bookmark or comment a marker belongs to: a bookmark by name, a comment by id.</summary>
    private static string MarkerIdentity(XElement marker)
    {
        var id = (string?)marker.Attribute(W.id) ?? string.Empty;
        if (marker.Name == W.bookmarkStart)
            return "b\u0000" + ((string?)marker.Attribute(W.name) ?? "#" + id);
        if (marker.Name == W.bookmarkEnd)
        {
            var names = marker.Document is { } document
                ? BookmarkNamesById.GetValue(document, d => d.Descendants(W.bookmarkStart)
                    .Where(s => s.Attribute(W.id) is not null && s.Attribute(W.name) is not null)
                    .GroupBy(s => (string)s.Attribute(W.id)!)
                    .ToDictionary(g => g.Key, g => (string)g.First().Attribute(W.name)!, StringComparer.Ordinal))
                : null;
            return "b\u0000" + (names?.GetValueOrDefault(id) ?? "#" + id);
        }
        return "c\u0000" + id;
    }

    /// <summary>Each source document's bookmark names by id, built once per document: an end names its bookmark
    /// only through its start. The renderer only reads its source documents.</summary>
    private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<XDocument, Dictionary<string, string>>
        BookmarkNamesById = new();
}
