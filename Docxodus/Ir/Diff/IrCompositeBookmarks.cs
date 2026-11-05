using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Docxodus.Ir.Diff;

internal static partial class IrMarkupRenderer
{
    // Equal spans of a composed paragraph read the base. Carry each reviewer's marker changes
    // separately, in the base's rendering coordinates, so unchanged text retains those edits too.
    private sealed class ComposedMarkerPlan(IrDocument original, IReadOnlyList<IrDocument> reviewers)
    {
        public List<(int Position, List<SourceBoundaryStep> Steps)> Pending { get; } = new();

        public List<XElement> Slice(int start, int end, SourceRunModel runs, RenderState state)
        {
            var result = new List<XElement>();
            int cursor = start;
            foreach (var gap in Pending.Where(e => e.Position >= start && e.Position <= end)
                         .OrderBy(e => e.Position).ToList())
            {
                if (cursor < gap.Position) result.AddRange(runs.Slice(cursor, gap.Position));
                result.AddRange(EmitSourceBoundaries(new() { gap.Steps }, runs, runs, state, reviewers, original));
                cursor = gap.Position;
                Pending.Remove(gap);
            }
            result.AddRange(runs.Slice(cursor, end));
            return result;
        }
    }

    private static ComposedMarkerPlan? PlanComposedMarkers(IrCompositeOp op,
        IReadOnlyList<IrAuthoredTokenOp> authored, IReadOnlyList<IrDiffToken> leftTokens,
        IReadOnlyDictionary<int, SourceBoundarySource> sources, IrDocument original,
        IReadOnlyList<IrDocument> reviewers, RenderState state,
        ref List<List<SourceBoundaryStep>>[]? emitBefore)
    {
        if (!sources.Values.Any(source => source.Runs.HasRangeMarkers)) return null;
        var plan = new ComposedMarkerPlan(original, reviewers);
        var left = sources[-1].Runs;
        var leftKeys = left.MarkerKeysBetween(0, int.MaxValue);
        var anchors = (op.SourceRightAnchors ?? IrNodeList.Empty<IrSourceRightAnchor>()).ToDictionary(s => s.Reviewer);
        var copies = sources.ToDictionary(pair => pair.Key, pair =>
        {
            var available = pair.Value.Runs.UnrepresentedMarkers(0, int.MaxValue, new HashSet<string>())
                .SelectMany(group => group).ToHashSet();
            return pair.Value.Runs.StandaloneMarkers(Array.Empty<SourceRunModel.SourceBoundary>())
                .Where(boundary => available.Contains(boundary.Primary!)).ToList();
        });
        int? Map(int reviewer, int position, bool fromLeft)
        {
            if (reviewer < 0) return position;
            if (anchors[reviewer].TokenDiff is not { } diff) return null;
            foreach (var tokenOp in diff.Ops.Where(tokenOp => KeepsTextUnchanged(tokenOp, leftTokens, sources[reviewer].Tokens)))
            {
                var (ls, le) = LeftSpanPositions(leftTokens, tokenOp);
                var (rs, re) = RightSpanPositions(sources[reviewer].Tokens, tokenOp);
                if (fromLeft ? position >= ls && position <= le : position >= rs && position <= re)
                    return fromLeft ? rs + position - ls : ls + position - rs;
            }
            // Paragraph edges remain gaps even when the first or last word was replaced.
            int leftEnd = leftTokens.LastOrDefault()?.EndChar ?? 0;
            int rightEnd = sources[reviewer].Tokens.LastOrDefault()?.EndChar ?? 0;
            if (position == 0) return 0;
            if (position == (fromLeft ? leftEnd : rightEnd)) return fromLeft ? rightEnd : leftEnd;
            return null;
        }
        var positions = new SortedSet<int>();
        foreach (var (reviewer, source) in sources.Where(pair => pair.Key >= 0))
        {
            foreach (var group in left.UnrepresentedMarkers(0, int.MaxValue, source.Runs.MarkerKeysBetween(0, int.MaxValue)))
                positions.Add(group.Key);
            foreach (var group in source.Runs.UnrepresentedMarkers(0, int.MaxValue, leftKeys))
                if (Map(reviewer, group.Key, false) is { } position) positions.Add(position);
        }
        foreach (int position in positions)
        {
            var points = sources.Keys.ToDictionary(reviewer => reviewer, reviewer => Map(reviewer, position, true));
            // A changed endpoint inside text another reviewer replaced has no common source gap.
            // Leave such content to its owning revision; do not invent a location within the replacement.
            if (points.Values.Any(point => point is null)) continue;
            var baseTypes = copies[-1].Where(marker => marker.Position == position).Select(marker => marker.Type).ToHashSet();
            var allBaseTypes = copies[-1].Select(marker => marker.Type).ToHashSet();
            var boundaries = sources.ToDictionary(pair => pair.Key, pair =>
                new SourceRunModel.SourceBoundary("markers", points[pair.Key]!.Value,
                    copies[pair.Key].Where(boundary =>
                        (baseTypes.Contains(boundary.Type) || !allBaseTypes.Contains(boundary.Type) &&
                            boundary.Position == points[pair.Key] && Map(pair.Key, boundary.Position, false) == position) &&
                        (pair.Key < 0 || state.CompositeBookmarks?.Allows(pair.Key, boundary.Primary!) != false))
                        .SelectMany(boundary => boundary.Items).Select(item => item with { IsCore = false }).ToList(), -2, false));
            var steps = PlaceComposedSourceBoundary(boundaries[-1], left, boundaries, sources);
            if (!steps.Any(step => step.Kind is not null)) continue;
            var changedMembers = steps.Where(step => step.Kind is not null).SelectMany(step => step.Members).ToHashSet();
            var changedTypes = copies.Values.SelectMany(markers => markers)
                .Where(marker => marker.Members.Any(changedMembers.Contains)).Select(marker => marker.Type).ToHashSet();
            foreach (var (reviewer, boundary) in boundaries)
                sources[reviewer].Runs.WithholdSourceBoundary(boundary.Members.Concat(copies[reviewer]
                    .Where(marker => changedTypes.Contains(marker.Type)).SelectMany(marker => marker.Members)));
            // Inside a retained word, split its source slice at the exact character. At a token gap,
            // emit after every preceding deletion/insertion, including a replaced final word.
            int inside = -1, before = 0;
            for (int index = 0; index < authored.Count; index++)
            {
                var entry = authored[index];
                var (start, end) = entry.Op.Kind == IrTokenOpKind.Insert
                    ? RightSpanPositions(sources[entry.SourceReviewer].Tokens, entry.Op)
                    : LeftSpanPositions(leftTokens, entry.Op);
                int point = entry.Op.Kind == IrTokenOpKind.Insert ? points[entry.SourceReviewer]!.Value : position;
                if (end <= point) before = index + 1;
                if (entry.SourceReviewer < 0 && start < point && point < end) inside = index;
            }
            if (inside >= 0) plan.Pending.Add((position, steps));
            else
            {
                emitBefore ??= new List<List<SourceBoundaryStep>>[authored.Count + 1];
                (emitBefore[before] ??= new()).Add(steps);
            }
        }
        return plan.Pending.Count == 0 ? null : plan;
    }

    /// <summary>One main-story choice for a name added by several reviewers. Both block and inline
    /// emission consult it; choosing endpoints separately could splice two reviewers' ranges together.</summary>
    internal sealed class CompositeBookmarkPlan
    {
        private readonly Dictionary<(int Reviewer, string Id), bool> _choices = new();

        public void Choose(int reviewer, string id, bool selected) => _choices[(reviewer, id)] = selected;

        public bool Includes(int reviewer, XElement marker) =>
            (marker.Name == W.bookmarkStart || marker.Name == W.bookmarkEnd) &&
            _choices.ContainsKey((reviewer, (string?)marker.Attribute(W.id) ?? ""));

        public bool Allows(int reviewer, XElement marker) =>
            marker.Name != W.bookmarkStart && marker.Name != W.bookmarkEnd ||
            !_choices.TryGetValue((reviewer, (string?)marker.Attribute(W.id) ?? ""), out bool selected) || selected;
    }

    private sealed record BookmarkPoint(string Anchor, string Offset);

    private sealed record ReviewerBookmark(int Reviewer, string Id, string Name,
        BookmarkPoint Start, BookmarkPoint End, string Text);

    /// <summary>Record competing new names before rendering, using the pairwise alignments to compare
    /// ranges in the base's coordinates. Consensus keeps one attributed copy. StackAll also keeps the
    /// first candidate: one bookmark name cannot identify several destinations.</summary>
    internal static CompositeBookmarkPlan? PlanCompositeBookmarks(IrDocument baseIr,
        IReadOnlyList<(string Author, IrDocument Ir)> reviewers, IReadOnlyList<IrEditScript> scripts,
        ConflictResolution policy, IrDiffSettings settings, List<IrConflict> conflicts, ref int nextConflictId)
    {
        static XElement? Body(IrDocument ir) => ir.Body.PartUri is { } uri && ir.Sources.TryGetValue(uri, out var source)
            ? source.Root?.Element(W.body) : null;
        var baseNames = Body(baseIr)?.Descendants(W.bookmarkStart)
            .Select(e => (string?)e.Attribute(W.name)).OfType<string>().ToHashSet(StringComparer.Ordinal) ?? new();
        var sharedNames = reviewers.SelectMany((reviewer, index) => (Body(reviewer.Ir)?.Descendants(W.bookmarkStart)
                ?? Enumerable.Empty<XElement>()).Select(e => (Reviewer: index, Name: (string?)e.Attribute(W.name))))
            .Where(b => b.Name is not null && !baseNames.Contains(b.Name)).GroupBy(b => b.Name!, StringComparer.Ordinal)
            .Where(g => g.Select(b => b.Reviewer).Distinct().Count() > 1).Select(g => g.Key).ToHashSet(StringComparer.Ordinal);
        if (sharedNames.Count == 0) return null;
        var ranges = new List<ReviewerBookmark>();
        for (int r = 0; r < reviewers.Count; r++)
        {
            if (Body(reviewers[r].Ir) is not { } body || !body.Descendants(W.bookmarkStart)
                    .Any(e => sharedNames.Contains((string?)e.Attribute(W.name) ?? "")))
                continue;
            var map = BookmarkParagraphMap(baseIr, reviewers[r].Ir, scripts[r].Operations, settings);
            var models = new Dictionary<XElement, SourceRunModel>(ReferenceEqualityComparer.Instance);
            var tokens = new Dictionary<XElement, (IReadOnlyList<IrDiffToken> Left, IReadOnlyList<IrDiffToken> Right)>(ReferenceEqualityComparer.Instance);
            SourceRunModel Model(XElement p)
            {
                if (!models.TryGetValue(p, out var model)) models[p] = model = new SourceRunModel(p);
                return model;
            }
            var elements = body.Descendants().ToList();
            var ends = elements.Where(e => e.Name == W.bookmarkEnd).GroupBy(e => (string?)e.Attribute(W.id) ?? "")
                .ToDictionary(g => g.Key, g => g.First());
            foreach (var start in elements.Where(e => e.Name == W.bookmarkStart))
            {
                if ((string?)start.Attribute(W.name) is not { } name || !sharedNames.Contains(name) ||
                    (string?)start.Attribute(W.id) is not { } id || !ends.TryGetValue(id, out var end))
                    continue;
                ranges.Add(new ReviewerBookmark(r, id, name, Point(start, false), Point(end, true),
                    string.Concat(elements.SkipWhile(e => e != start).TakeWhile(e => e != end)
                        .Where(e => e.Name == W.t).Select(e => e.Value))));
            }

            BookmarkPoint Point(XElement marker, bool end)
            {
                var paragraph = marker.Ancestors(W.p).FirstOrDefault();
                int offset;
                if (paragraph is not null)
                    offset = Model(paragraph).MarkerOffset(marker) ?? 0;
                else
                {
                    int index = elements.IndexOf(marker);
                    paragraph = end ? elements.Take(index).LastOrDefault(e => e.Name == W.p)
                        : elements.Skip(index + 1).FirstOrDefault(e => e.Name == W.p);
                    offset = end && paragraph is not null ? Model(paragraph).TextLength : 0;
                }
                if (paragraph is null)
                    return new BookmarkPoint(baseIr.Body.Blocks.FirstOrDefault()?.Anchor.ToString() ?? "", "empty");
                if (!map.TryGetValue(paragraph, out var pair))
                    return new BookmarkPoint(baseIr.Body.Blocks.FirstOrDefault()?.Anchor.ToString() ?? "",
                        "insert:" + string.Join("/", paragraph.AncestorsAndSelf()
                        .TakeWhile(e => e != body).Reverse().Select(e => e.Name.LocalName + ":" +
                            e.ElementsBeforeSelf().Count(c => !IsBlockLevelMarker(c)))) + ":" +
                        string.Concat(paragraph.Descendants(W.t).Select(t => t.Value)) + "@" + offset);
                if (pair.Op?.TokenDiff is not { } diff || pair.Left is not IrParagraph left || pair.Right is not IrParagraph right)
                    return new BookmarkPoint(pair.Anchor, offset.ToString());
                if (!tokens.TryGetValue(paragraph, out var sourceTokens))
                    tokens[paragraph] = sourceTokens = (IrDiffTokenizer.Tokenize(left, settings), IrDiffTokenizer.Tokenize(right, settings));
                var (lt, rt) = sourceTokens;
                int At(IReadOnlyList<IrDiffToken> tokens, int i) => i < tokens.Count ? tokens[i].StartChar
                    : tokens.LastOrDefault()?.EndChar ?? 0;
                // Equal text translates exactly, including a bookmark at either endpoint. A position inside
                // inserted text also includes that text, so different insertions at one base gap do not agree.
                foreach (var op in diff.Ops.OrderBy(o => o.Kind == IrTokenOpKind.Insert ? 1 : 0))
                {
                    int start = At(rt, op.RightStart), stop = At(rt, op.RightEnd);
                    if (op.RightLength == 0 || offset < start || offset > stop) continue;
                    string mapped = op.Kind == IrTokenOpKind.Insert
                        ? $"{At(lt, op.LeftStart)}+{string.Concat(rt.Skip(op.RightStart).Take(op.RightLength).Select(t => t.Text))}@{offset - start}"
                        : (At(lt, op.LeftStart) + offset - start).ToString();
                    return new BookmarkPoint(pair.Anchor, mapped);
                }
                return new BookmarkPoint(pair.Anchor, offset.ToString());
            }
        }

        CompositeBookmarkPlan? plan = null;
        foreach (var group in ranges.GroupBy(b => b.Name, StringComparer.Ordinal).Where(g => g.Select(b => b.Reviewer).Distinct().Count() > 1))
        {
            plan ??= new CompositeBookmarkPlan();
            var first = group.First();
            bool contested = group.Any(b => b.Start != first.Start || b.End != first.End);
            if (contested)
                conflicts.Add(new IrConflict(nextConflictId++, first.Start.Anchor, 0, 0, policy,
                    IrNodeList.From(group.Select(b => new IrConflictCompetitor(reviewers[b.Reviewer].Author,
                        $"{b.Name}: {b.Text} [{b.Start.Anchor}@{b.Start.Offset}..{b.End.Anchor}@{b.End.Offset}]")))));
            int winner = contested && policy == ConflictResolution.BaseWins ? -1 : first.Reviewer;
            foreach (var range in group)
                plan.Choose(range.Reviewer, range.Id, range.Reviewer == winner);
        }
        return plan;
    }

    internal static void ReconcileCompositeBookmarkClones(RenderState state, IReadOnlyList<string> authors)
    {
        if (state.CompositeBookmarks is not { } plan) return;
        var savedAuthor = state.AuthorOverride;
        foreach (var (marker, reviewer) in state.CompositeBookmarkClones)
        {
            if (marker.Parent is null) continue;
            if (!plan.Allows(reviewer, marker)) RemoveBookmarkMarker(marker);
            else if (!IsTrackedBookmarkMarker(marker))
            {
                state.AuthorOverride = authors[reviewer];
                WrapMarkerInRevision(marker, RevKind.Ins, state);
            }
        }
        state.AuthorOverride = savedAuthor;
    }

    private sealed record BookmarkParagraphPair(string Anchor, IrBlock? Left, IrBlock? Right, IrEditOp? Op);

    private static Dictionary<XElement, BookmarkParagraphPair> BookmarkParagraphMap(IrDocument left, IrDocument right,
        IReadOnlyList<IrEditOp> operations, IrDiffSettings settings)
    {
        var result = new Dictionary<XElement, BookmarkParagraphPair>(ReferenceEqualityComparer.Instance);
        var leftParagraphs = new Dictionary<XElement, IrParagraph>(ReferenceEqualityComparer.Instance);
        foreach (var paragraph in left.AnchorIndex.Values.OfType<IrParagraph>())
            if (paragraph.Source.Element is { } source) leftParagraphs[source] = paragraph;
        void PairEqual(XElement? l, XElement? r)
        {
            if (l is null || r is null) return;
            var lp = l.DescendantsAndSelf(W.p).ToList();
            var rp = r.DescendantsAndSelf(W.p).ToList();
            if (lp.Count != rp.Count) return;
            for (int i = 0; i < lp.Count; i++)
                if (leftParagraphs.TryGetValue(lp[i], out var paragraph))
                    result[rp[i]] = new BookmarkParagraphPair(paragraph.Anchor.ToString(), paragraph, null, null);
        }
        void Walk(IReadOnlyList<IrEditOp> ops)
        {
            foreach (var op in ops)
            {
                var l = op.LeftAnchor is { } la ? left.AnchorIndex.GetValueOrDefault(la) : null;
                var r = op.RightAnchor is { } ra ? right.AnchorIndex.GetValueOrDefault(ra) : null;
                if (l is IrParagraph && r is IrParagraph && r.Source.Element is { } p)
                    result[p] = new BookmarkParagraphPair(l.Anchor.ToString(), l, r, op);
                else if (op.Kind is IrEditOpKind.EqualBlock or IrEditOpKind.FormatOnlyBlock ||
                         IrCompositeMerger.MarkerOnlyWrapper(op, left, right, settings))
                    PairEqual(l?.Source.Element, r?.Source.Element);
                if (op.WrapperOps is { } children) Walk(children);
                if (op.TableDiff is not { } table) continue;
                foreach (var row in table.RowOps)
                {
                    if (row.Kind == IrRowOpKind.EqualRow)
                        PairEqual(row.LeftRowAnchor is { } lr ? FindRowSource(left, lr) : null,
                            row.RightRowAnchor is { } rr ? FindRowSource(right, rr) : null);
                    if (row.CellOps is not { } cells) continue;
                    foreach (var cell in cells)
                        if (cell.BlockOps is { } blocks) Walk(blocks);
                        else PairEqual(cell.LeftCellAnchor is { } lc ? FindCellSource(left, lc) : null,
                            cell.RightCellAnchor is { } rc ? FindCellSource(right, rc) : null);
                }
            }
        }
        Walk(operations);
        return result;
    }
}
