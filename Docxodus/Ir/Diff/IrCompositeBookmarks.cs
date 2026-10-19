using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Docxodus.Ir.Diff;

internal static partial class IrMarkupRenderer
{
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
