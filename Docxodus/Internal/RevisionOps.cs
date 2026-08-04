#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Docxodus.Internal;

/// <summary>
/// Markup-native tracked-revision enumeration and selective per-revision resolution
/// (issue #318). Reads <c>w:ins</c>/<c>w:del</c>/<c>w:moveFrom</c>/<c>w:moveTo</c> and the
/// <c>*PrChange</c> format-change family directly off the live session XML — no
/// accept-all/reject-all re-diff — grouping physically contiguous markup of the same
/// kind and author into one revision per user-visible change, keyed by a stable id
/// derived from the markup's own <c>w:id</c> attributes. <see cref="Apply"/> resolves
/// exactly one group in place, mirroring <see cref="RevisionProcessor"/>'s per-element
/// semantics (unwrap vs. remove, <c>w:delText</c> restore, paragraph-mark coalescing
/// into the following paragraph, row removal, stored-property restore).
///
/// v1 scope: run-content ins/del (any story), paragraph-mark ins/del, table-row
/// ins/del (<c>w:trPr</c> markers absorb their row's content markup), named move
/// pairs (both sides resolve together), and the format-change family
/// (<c>rPrChange</c>/<c>pPrChange</c>/<c>sectPrChange</c>/<c>tblPrChange</c>/
/// <c>trPrChange</c>/<c>tcPrChange</c>/<c>tblGridChange</c>/<c>tblPrExChange</c>).
/// Exotic families without per-revision semantics here (<c>cellIns</c>/<c>cellDel</c>/
/// <c>cellMerge</c>, content-control ins/del ranges, <c>numPr/ins</c>) are not
/// enumerated; whole-document accept-all/reject-all still handles them.
/// </summary>
internal static class RevisionOps
{
    internal const string TypeInsert = "insert";
    internal const string TypeDelete = "delete";
    internal const string TypeMove = "move";
    internal const string TypeFormat = "format";

    internal enum UnitKind { Content, ParaMark, RowMark, PropsChange }

    /// <summary>One revision markup element, positioned in document order (a paragraph's
    /// mark unit is repositioned to the END of its paragraph — that is where the pilcrow
    /// lives semantically, and what makes multi-paragraph runs of markup group).</summary>
    internal sealed class RevisionUnit
    {
        public required XElement Element { get; init; }
        public required UnitKind Kind { get; init; }
        public required string Type { get; init; }
        public required string Author { get; init; }
        public string? Date { get; init; }
        /// <summary>Move-range name when the unit sits inside a named move range — such
        /// units group per name (both sides of the pair) rather than by adjacency.</summary>
        public string? MoveName { get; init; }
        public XElement? Paragraph { get; init; }
        /// <summary>For RowMark: the <c>w:tr</c> itself. For Content: the marked row the
        /// unit sits inside (so the row group absorbs it), else null.</summary>
        public XElement? MarkedRow { get; init; }
        public long? Wid { get; init; }
    }

    internal sealed class RevisionGroup
    {
        public string Id { get; set; } = "";
        public required string Type { get; init; }
        public required string Author { get; init; }
        public string? Date { get; set; }
        public required int PartIndex { get; init; }
        public List<RevisionUnit> Units { get; } = new();
        /// <summary>Move-range marker elements (start/end, both sides) removed when the
        /// group resolves.</summary>
        public List<XElement> RangeMarkers { get; } = new();
    }

    /// <summary>
    /// The live XML boundaries a native comment can bracket for one revision group.
    /// <see cref="First"/>/<see cref="Last"/> are inclusive element boundaries;
    /// when a revision has no commentable inline content (for example, a paragraph-mark
    /// revision), <see cref="PointParagraph"/> supplies a legal collapsed anchor.
    /// </summary>
    internal sealed class RevisionCommentTarget
    {
        public XElement? First { get; init; }
        public XElement? Last { get; init; }
        public XElement? PointParagraph { get; init; }
    }

    private static readonly XName[] RevWrapperNames = { W.ins, W.del, W.moveFrom, W.moveTo };

    private static readonly HashSet<XName> PropsChangeNames = new()
    {
        W.rPrChange, W.pPrChange, W.sectPrChange, W.tblPrChange,
        W.trPrChange, W.tcPrChange, W.tblGridChange, W.tblPrExChange,
    };

    // Elements transparent to adjacency: two revision units with only these between
    // them belong to the same user-visible change.
    private static readonly HashSet<XName> IgnorableBetween = new()
    {
        W.proofErr, W.bookmarkStart, W.bookmarkEnd, W.commentRangeStart, W.commentRangeEnd,
        W.moveFromRangeStart, W.moveFromRangeEnd, W.moveToRangeStart, W.moveToRangeEnd, W.pPr,
    };

    // ─── Enumeration ────────────────────────────────────────────────────

    internal static List<RevisionGroup> Enumerate(IReadOnlyList<XElement> partRoots)
    {
        var groups = new List<RevisionGroup>();
        for (int pi = 0; pi < partRoots.Count; pi++)
        {
            var ctx = new WalkCtx();
            var units = new List<RevisionUnit>();
            WalkChildren(partRoots[pi].Elements(), ctx, null, null, null, units);
            BuildGroups(units, ctx, pi, groups);
        }
        AssignIds(groups);
        return groups;
    }

    /// <summary>
    /// Resolve the exact live extent a Word comment should bracket for a revision. Content
    /// revisions bracket their outer revision wrappers, which keeps the comment markers outside
    /// markup that accept/reject may unwrap or remove. Run-format revisions bracket the affected
    /// runs. A move targets its destination (the proposed location); rejecting it therefore
    /// collapses the range at that location while accepting it leaves the moved text selected.
    /// Structural/paragraph-mark revisions fall back to the affected paragraph as a point target.
    /// </summary>
    internal static RevisionCommentTarget? CommentTarget(RevisionGroup group)
    {
        var content = group.Units
            .Where(u => u.Kind == UnitKind.Content
                && (group.Type != TypeMove || u.Element.Name == W.moveTo))
            .Select(u => u.Element)
            .Where(e => !group.Units.Any(u => u.Kind == UnitKind.Content
                && u.Element != e && e.Ancestors().Contains(u.Element)
                && (group.Type != TypeMove || u.Element.Name == W.moveTo)))
            .ToList();
        if (content.Count > 0)
            return new RevisionCommentTarget { First = content[0], Last = content[^1] };

        var formatRuns = group.Units
            .Where(u => u.Kind == UnitKind.PropsChange && u.Element.Name == W.rPrChange)
            .Select(u => u.Element.Parent?.Parent)
            .Where(r => r is not null && r.Name == W.r)
            .Select(r => r!)
            .Distinct()
            .ToList();
        if (formatRuns.Count > 0)
            return new RevisionCommentTarget { First = formatRuns[0], Last = formatRuns[^1] };

        // Non-run property changes affect their containing paragraph as a whole.
        var propertyParagraphs = group.Units
            .Where(u => u.Kind == UnitKind.PropsChange)
            .Select(AffectedParagraph)
            .Where(p => p is not null)
            .Select(p => p!)
            .Distinct()
            .ToList();
        if (propertyParagraphs.Count > 0)
        {
            var firstInline = propertyParagraphs[0].Elements().FirstOrDefault(e => e.Name != W.pPr);
            var lastInline = propertyParagraphs[^1].Elements().LastOrDefault(e => e.Name != W.pPr);
            if (firstInline is not null && lastInline is not null)
                return new RevisionCommentTarget { First = firstInline, Last = lastInline };
            return new RevisionCommentTarget { PointParagraph = propertyParagraphs[0] };
        }

        // Paragraph/row marks have no textual extent. Anchor them as a point in the affected
        // paragraph; if resolution removes that paragraph, the existing merge path carries the
        // point into the surviving paragraph.
        var pointParagraph = group.Units
            .Select(AffectedParagraph)
            .FirstOrDefault(p => p is not null);
        return pointParagraph is null
            ? null
            : new RevisionCommentTarget { PointParagraph = pointParagraph };
    }

    private static XElement? AffectedParagraph(RevisionUnit unit)
    {
        if (unit.Paragraph is not null) return unit.Paragraph;
        if (unit.MarkedRow?.Descendants(W.p).FirstOrDefault() is { } rowParagraph)
            return rowParagraph;
        if (unit.Element.Ancestors(W.p).FirstOrDefault() is { } ancestorParagraph)
            return ancestorParagraph;

        // Table/cell property changes live above their text-bearing paragraphs. A comment
        // reference still needs a paragraph host, so use the first paragraph in that owner.
        var tableOwner = unit.Element.Ancestors()
            .FirstOrDefault(e => e.Name == W.tc || e.Name == W.tr || e.Name == W.tbl);
        if (tableOwner?.Descendants(W.p).FirstOrDefault() is { } tableParagraph)
            return tableParagraph;

        // A final body-level sectPr has no paragraph ancestor. Its nearest legal review anchor
        // is the preceding body paragraph (an in-paragraph section break was handled above).
        return unit.Element.Ancestors(W.body).FirstOrDefault()?.Elements(W.p).LastOrDefault();
    }

    private sealed class WalkCtx
    {
        public readonly List<(string? Id, string Name)> MoveFromStack = new();
        public readonly List<(string? Id, string Name)> MoveToStack = new();
        public readonly Dictionary<string, List<XElement>> RangeMarkers = new(StringComparer.Ordinal);

        public List<XElement> MarkersFor(string name)
        {
            if (!RangeMarkers.TryGetValue(name, out var list))
                RangeMarkers[name] = list = new List<XElement>();
            return list;
        }
    }

    private static void WalkChildren(IEnumerable<XElement> children, WalkCtx ctx,
        XElement? paragraph, XElement? markedRow, string? markedRowType, List<RevisionUnit> sink)
    {
        foreach (var child in children)
        {
            var n = child.Name;
            if (n == W.moveFromRangeStart || n == W.moveToRangeStart)
            {
                var name = (string?)child.Attribute(W.name);
                if (name is not null)
                {
                    var stack = n == W.moveFromRangeStart ? ctx.MoveFromStack : ctx.MoveToStack;
                    stack.Add(((string?)child.Attribute(W.id), name));
                    ctx.MarkersFor(name).Add(child);
                }
                continue;
            }
            if (n == W.moveFromRangeEnd || n == W.moveToRangeEnd)
            {
                var stack = n == W.moveFromRangeEnd ? ctx.MoveFromStack : ctx.MoveToStack;
                var id = (string?)child.Attribute(W.id);
                int idx = stack.FindLastIndex(e => e.Id == id);
                if (idx < 0) idx = stack.Count - 1;
                if (idx >= 0)
                {
                    ctx.MarkersFor(stack[idx].Name).Add(child);
                    stack.RemoveAt(idx);
                }
                continue;
            }
            if (n == W.p) { WalkParagraph(child, ctx, markedRow, markedRowType, sink); continue; }
            if (n == W.tr) { WalkRow(child, ctx, sink); continue; }
            if ((n == W.ins || n == W.del || n == W.moveFrom || n == W.moveTo) && IsContentWrapper(child))
            {
                sink.Add(MakeUnit(child, UnitKind.Content, paragraph, markedRow, markedRowType, ctx));
                WalkChildren(child.Elements(), ctx, paragraph, markedRow, markedRowType, sink);
                continue;
            }
            if (PropsChangeNames.Contains(n))
            {
                sink.Add(MakePropsUnit(child, paragraph));
                continue;
            }
            // pPr/trPr are handled by WalkParagraph/WalkRow; everything else (runs,
            // hyperlinks, sdt, tbl, tc, rPr, tblPr, tblGrid, sectPr, …) recurses so
            // nested wrappers and *PrChange elements are found wherever they live.
            if (n == W.pPr || n == W.trPr) continue;
            if (child.HasElements)
                WalkChildren(child.Elements(), ctx, paragraph, markedRow, markedRowType, sink);
        }
    }

    private static void WalkParagraph(XElement p, WalkCtx ctx,
        XElement? markedRow, string? markedRowType, List<RevisionUnit> sink)
    {
        var pPr = p.Element(W.pPr);
        if (pPr is not null)
        {
            foreach (var pc in pPr.Descendants().Where(d => PropsChangeNames.Contains(d.Name)))
                sink.Add(MakePropsUnit(pc, p));
        }

        WalkChildren(p.Elements().Where(e => e.Name != W.pPr), ctx, p, markedRow, markedRowType, sink);

        // The paragraph-mark revision is emitted LAST: the pilcrow sits at the end of the
        // paragraph, which is what lets a fully inserted/deleted paragraph — runs plus
        // mark plus the next paragraph's runs — group into one revision.
        var mark = pPr?.Element(W.rPr)?.Elements()
            .FirstOrDefault(e => RevWrapperNames.Contains(e.Name));
        if (mark is not null)
            sink.Add(MakeUnit(mark, UnitKind.ParaMark, p, markedRow, markedRowType, ctx));
    }

    private static void WalkRow(XElement tr, WalkCtx ctx, List<RevisionUnit> sink)
    {
        var trPr = tr.Element(W.trPr);
        string? rowType = null;
        if (trPr is not null)
        {
            foreach (var pc in trPr.Descendants().Where(d => PropsChangeNames.Contains(d.Name)))
                sink.Add(MakePropsUnit(pc, null));
            var mark = trPr.Elements().FirstOrDefault(e => e.Name == W.ins || e.Name == W.del);
            if (mark is not null)
            {
                rowType = mark.Name == W.ins ? TypeInsert : TypeDelete;
                sink.Add(new RevisionUnit
                {
                    Element = mark,
                    Kind = UnitKind.RowMark,
                    Type = rowType,
                    Author = AuthorOf(mark),
                    Date = (string?)mark.Attribute(W.date),
                    MarkedRow = tr,
                    Wid = WidOf(mark),
                });
            }
        }
        WalkChildren(tr.Elements().Where(e => e.Name != W.trPr), ctx,
            null, rowType is not null ? tr : null, rowType, sink);
    }

    /// <summary>A revision wrapper is CONTENT when its parent isn't a property container —
    /// <c>w:rPr</c> holds paragraph-mark revisions, <c>w:trPr</c> row marks,
    /// <c>w:numPr</c>/<c>m:ctrlPr</c> revision flavors this v1 does not enumerate.</summary>
    private static bool IsContentWrapper(XElement el)
    {
        var pn = el.Parent?.Name;
        return pn != W.rPr && pn != W.trPr && pn != W.numPr && pn != M.ctrlPr;
    }

    private static RevisionUnit MakeUnit(XElement el, UnitKind kind, XElement? paragraph,
        XElement? markedRow, string? markedRowType, WalkCtx ctx)
    {
        var n = el.Name;
        string type = n == W.ins ? TypeInsert : n == W.del ? TypeDelete : TypeMove;
        string? moveName = null;
        if (n == W.moveFrom && ctx.MoveFromStack.Count > 0) moveName = ctx.MoveFromStack[^1].Name;
        else if (n == W.moveTo && ctx.MoveToStack.Count > 0) moveName = ctx.MoveToStack[^1].Name;
        return new RevisionUnit
        {
            Element = el,
            Kind = kind,
            Type = type,
            Author = AuthorOf(el),
            Date = (string?)el.Attribute(W.date),
            MoveName = moveName,
            Paragraph = paragraph,
            MarkedRow = markedRowType == type ? markedRow : null,
            Wid = WidOf(el),
        };
    }

    private static RevisionUnit MakePropsUnit(XElement el, XElement? paragraph) =>
        new()
        {
            Element = el,
            Kind = UnitKind.PropsChange,
            Type = TypeFormat,
            Author = AuthorOf(el),
            Date = (string?)el.Attribute(W.date),
            Paragraph = paragraph,
            Wid = WidOf(el),
        };

    private static string AuthorOf(XElement el) => (string?)el.Attribute(W.author) ?? "unknown";

    private static long? WidOf(XElement el) =>
        long.TryParse((string?)el.Attribute(W.id), out var v) ? v : null;

    // ─── Grouping ───────────────────────────────────────────────────────

    private static void BuildGroups(List<RevisionUnit> units, WalkCtx ctx, int partIndex,
        List<RevisionGroup> groups)
    {
        var moveGroups = new Dictionary<string, RevisionGroup>(StringComparer.Ordinal);
        var rowGroupByTr = new Dictionary<XElement, RevisionGroup>();
        RevisionGroup? cur = null;
        RevisionGroup? lastRowGroup = null;

        foreach (var u in units)
        {
            if (u.MoveName is not null)
            {
                if (!moveGroups.TryGetValue(u.MoveName, out var mg))
                {
                    mg = new RevisionGroup { Type = TypeMove, Author = u.Author, Date = u.Date, PartIndex = partIndex };
                    if (ctx.RangeMarkers.TryGetValue(u.MoveName, out var markers))
                        mg.RangeMarkers.AddRange(markers);
                    moveGroups[u.MoveName] = mg;
                    groups.Add(mg);
                }
                mg.Units.Add(u);
                continue;
            }

            if (u.Kind == UnitKind.RowMark)
            {
                if (lastRowGroup is not null && lastRowGroup.Type == u.Type && lastRowGroup.Author == u.Author
                    && lastRowGroup.Units[^1].MarkedRow is { } prevTr && u.MarkedRow is { } tr2
                    && prevTr.Parent == tr2.Parent && OnlyIgnorableBetween(prevTr, tr2))
                {
                    lastRowGroup.Units.Add(u);
                }
                else
                {
                    lastRowGroup = NewGroup(u, partIndex);
                    groups.Add(lastRowGroup);
                }
                rowGroupByTr[u.MarkedRow!] = lastRowGroup;
                cur = null;
                continue;
            }

            if ((u.Kind == UnitKind.Content || u.Kind == UnitKind.ParaMark) && u.MarkedRow is not null
                && rowGroupByTr.TryGetValue(u.MarkedRow, out var hostRow) && hostRow.Type == u.Type)
            {
                hostRow.Units.Add(u);
                continue;
            }

            if (u.Kind == UnitKind.PropsChange)
            {
                if (cur is not null && cur.Type == TypeFormat && cur.Author == u.Author
                    && cur.Date == u.Date
                    && cur.Units[^1].Element.Name == W.rPrChange && u.Element.Name == W.rPrChange
                    && AdjacentFormatRuns(cur.Units[^1].Element, u.Element))
                {
                    cur.Units.Add(u);
                }
                else
                {
                    cur = NewGroup(u, partIndex);
                    groups.Add(cur);
                }
                continue;
            }

            // Content / paragraph-mark insert, delete, or unranged move — adjacency grouping.
            if (cur is not null && cur.Type == u.Type && cur.Author == u.Author
                && cur.Units[^1].Element.Name == u.Element.Name
                && Contiguous(cur.Units[^1], u))
            {
                cur.Units.Add(u);
            }
            else
            {
                cur = NewGroup(u, partIndex);
                groups.Add(cur);
            }
        }
    }

    private static RevisionGroup NewGroup(RevisionUnit u, int partIndex)
    {
        var g = new RevisionGroup { Type = u.Type, Author = u.Author, Date = u.Date, PartIndex = partIndex };
        g.Units.Add(u);
        return g;
    }

    private static bool AdjacentFormatRuns(XElement prevChange, XElement curChange)
    {
        var runA = prevChange.Parent?.Parent;
        var runB = curChange.Parent?.Parent;
        return runA is not null && runB is not null
            && runA.Name == W.r && runB.Name == W.r
            && runA.Parent == runB.Parent && OnlyIgnorableBetween(runA, runB);
    }

    private static bool Contiguous(RevisionUnit prev, RevisionUnit cur)
    {
        if (prev.Kind == UnitKind.Content && cur.Kind == UnitKind.Content)
        {
            if (prev.Paragraph is null || cur.Paragraph is null)
                return prev.Element.Parent == cur.Element.Parent
                    && OnlyIgnorableBetween(prev.Element, cur.Element);
            if (prev.Paragraph != cur.Paragraph) return false;
            var a = TopLevelWithin(cur.Paragraph, prev.Element);
            var b = TopLevelWithin(cur.Paragraph, cur.Element);
            if (a is null || b is null) return false;
            if (a == b) return true; // both nested inside the same container (e.g. one hyperlink)
            return OnlyIgnorableBetween(a, b);
        }
        if (prev.Kind == UnitKind.Content && cur.Kind == UnitKind.ParaMark)
        {
            if (prev.Paragraph != cur.Paragraph || cur.Paragraph is null) return false;
            var a = TopLevelWithin(cur.Paragraph, prev.Element);
            return a is not null && a.ElementsAfterSelf().All(IsIgnorableBetween);
        }
        if (prev.Kind == UnitKind.ParaMark && cur.Kind == UnitKind.Content)
        {
            if (prev.Paragraph is null || cur.Paragraph is null) return false;
            if (!IsNextParagraph(prev.Paragraph, cur.Paragraph)) return false;
            var b = TopLevelWithin(cur.Paragraph, cur.Element);
            return b is not null && b.ElementsBeforeSelf().All(IsIgnorableBetween);
        }
        if (prev.Kind == UnitKind.ParaMark && cur.Kind == UnitKind.ParaMark)
        {
            if (prev.Paragraph is null || cur.Paragraph is null) return false;
            if (!IsNextParagraph(prev.Paragraph, cur.Paragraph)) return false;
            // An empty paragraph whose only substance is its (revised) mark.
            return cur.Paragraph.Elements().All(IsIgnorableBetween);
        }
        return false;
    }

    private static XElement? TopLevelWithin(XElement container, XElement el) =>
        el.AncestorsAndSelf().FirstOrDefault(x => x.Parent == container);

    private static bool IsNextParagraph(XElement p, XElement q)
    {
        foreach (var e in p.ElementsAfterSelf())
        {
            if (e == q) return true;
            if (!IsIgnorableBetween(e)) return false;
        }
        return false;
    }

    private static bool OnlyIgnorableBetween(XElement a, XElement b)
    {
        foreach (var e in a.ElementsAfterSelf())
        {
            if (e == b) return true;
            if (!IsIgnorableBetween(e)) return false;
        }
        return false;
    }

    private static bool IsIgnorableBetween(XElement element) =>
        IgnorableBetween.Contains(element.Name)
        || (element.Name == W.r && element.Descendants(W.commentReference).Any()
            && !element.Descendants(W.t).Any() && !element.Descendants(W.delText).Any());

    private static void AssignIds(List<RevisionGroup> groups)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        int fallback = 0;
        foreach (var g in groups)
        {
            long? min = null;
            foreach (var u in g.Units)
                if (u.Wid is { } w && (min is null || w < min)) min = w;
            foreach (var m in g.RangeMarkers)
                if (WidOf(m) is { } w && (min is null || w < min)) min = w;
            var baseId = min is { } mv
                ? "rev" + mv.ToString(System.Globalization.CultureInfo.InvariantCulture)
                : "revu" + fallback++;
            var id = baseId;
            int suffix = 2;
            while (!seen.Add(id)) id = baseId + "-" + suffix++;
            g.Id = id;
        }
    }

    // ─── Listing text ───────────────────────────────────────────────────

    internal static string GroupText(RevisionGroup g)
    {
        var sb = new StringBuilder();
        // A move pair carries the same text on both sides; render the source side only.
        bool moveFromOnly = g.Type == TypeMove
            && g.Units.Any(u => u.Element.Name == W.moveFrom);
        foreach (var u in g.Units)
        {
            if (moveFromOnly && u.Element.Name != W.moveFrom) continue;
            switch (u.Kind)
            {
                case UnitKind.Content:
                    AppendVisibleText(u.Element, u.Element.Name == W.del ? W.delText : W.t, sb);
                    break;
                case UnitKind.ParaMark:
                    sb.Append('¶');
                    break;
                case UnitKind.PropsChange:
                    if (u.Element.Name == W.rPrChange
                        && u.Element.Parent?.Parent is { } run && run.Name == W.r)
                        AppendVisibleText(run, W.t, sb);
                    else if (u.Element.Name == W.pPrChange
                        && u.Element.Parent?.Parent is { } para && para.Name == W.p)
                        AppendVisibleText(para, W.t, sb);
                    break;
            }
        }
        return sb.ToString();
    }

    private static void AppendVisibleText(XElement el, XName textName, StringBuilder sb)
    {
        foreach (var child in el.Elements())
        {
            var n = child.Name;
            if (n == W.pPr || n == W.rPr || n == W.trPr || n == W.tcPr || n == W.tblPr) continue;
            if (n == textName) { sb.Append(child.Value); continue; }
            if (n == W.tab) { sb.Append('\t'); continue; }
            if (n == W.br || n == W.cr) { sb.Append('\n'); continue; }
            if (child.HasElements) AppendVisibleText(child, textName, sb);
        }
    }

    // ─── Resolution ─────────────────────────────────────────────────────

    /// <summary>
    /// Resolve one revision group in place on the live tree. <paramref name="accept"/> keeps
    /// the change (insertions survive, deletions are carried out); reject restores the
    /// original (insertions are removed, deletions un-deleted). Returns the block-level
    /// elements (<c>w:p</c>/<c>w:tr</c>/<c>w:tbl</c>) the resolution detached, so the caller
    /// can report their anchors as removed.
    /// </summary>
    internal static List<XElement> Apply(RevisionGroup g, bool accept)
    {
        var removedBlocks = new List<XElement>();
        var touchedParagraphs = new HashSet<XElement>();

        var removedRows = g.Units
            .Where(u => u.Kind == UnitKind.RowMark
                && (u.Type == TypeInsert ? !accept : accept))
            .Select(u => u.MarkedRow!)
            .Distinct()
            .ToList();
        CollapseCommentsFromRemovedRows(removedRows);

        foreach (var u in g.Units.Where(u => u.Kind == UnitKind.PropsChange))
        {
            if (Detached(u.Element)) continue;
            if (accept) AcceptProps(u.Element);
            else RejectProps(u.Element);
        }

        foreach (var u in g.Units.Where(u => u.Kind == UnitKind.Content))
        {
            if (Detached(u.Element)) continue;
            if (u.Paragraph is not null) touchedParagraphs.Add(u.Paragraph);
            if (ContentSurvives(u.Element.Name, accept))
                UnwrapWrapper(u.Element, restoreDeleted: u.Element.Name == W.del);
            else
                RemoveContentWrapper(u.Element);
        }

        foreach (var u in g.Units.Where(u => u.Kind == UnitKind.RowMark))
        {
            if (Detached(u.Element)) continue;
            bool rowSurvives = u.Type == TypeInsert ? accept : !accept;
            if (rowSurvives) StripRowMark(u.Element);
            else RemoveRow(u.MarkedRow!, removedBlocks);
        }

        // Paragraph marks last, in reverse document order, so multi-paragraph coalescing
        // cascades into the single surviving paragraph exactly as RevisionProcessor's
        // grouped transform does.
        foreach (var u in g.Units.Where(u => u.Kind == UnitKind.ParaMark).Reverse())
        {
            if (Detached(u.Element) || u.Paragraph is null || Detached(u.Paragraph)) continue;
            touchedParagraphs.Add(u.Paragraph);
            if (ContentSurvives(u.Element.Name, accept))
            {
                u.Element.Remove();
            }
            else
            {
                var removedP = MergeParagraphWithNext(u.Paragraph);
                if (removedP is not null)
                {
                    removedBlocks.Add(removedP);
                    touchedParagraphs.Remove(removedP);
                }
                else
                {
                    // No following paragraph to coalesce into (last block of its
                    // container) — the mark cannot go away; strip the revision.
                    u.Element.Remove();
                }
            }
        }

        foreach (var m in g.RangeMarkers)
            if (!Detached(m)) m.Remove();

        foreach (var p in touchedParagraphs)
        {
            if (Detached(p)) continue;
            CleanParagraphHusks(p);
        }

        return removedBlocks;
    }

    /// <summary>
    /// A comment wholly contained by a row would otherwise lose its range/reference when
    /// selective resolution removes that row. Move complete marker triples to the nearest
    /// surviving paragraph as collapsed points before detaching the rows. This also protects
    /// comments authored in Word, not only comments created through the revision-target API.
    /// </summary>
    private static void CollapseCommentsFromRemovedRows(IReadOnlyCollection<XElement> removedRows)
    {
        if (removedRows.Count == 0) return;

        var removedSet = removedRows.ToHashSet();
        var root = removedRows.First().Document?.Root;
        if (root is null) return;

        var pivot = removedRows.First();
        var candidates = root.Descendants(W.p)
            .Where(p => !p.Ancestors(W.tr).Any(removedSet.Contains))
            .ToList();
        var host = candidates.FirstOrDefault(p => XNode.DocumentOrderComparer.Compare(p, pivot) > 0)
            ?? candidates.LastOrDefault(p => XNode.DocumentOrderComparer.Compare(p, pivot) < 0);
        if (host is null) return;

        var starts = removedRows.SelectMany(r => r.Descendants(W.commentRangeStart)).ToList();
        var ends = removedRows.SelectMany(r => r.Descendants(W.commentRangeEnd)).ToList();
        var referenceRuns = removedRows.SelectMany(r => r.Descendants(W.r))
            .Where(r => r.Descendants(W.commentReference).Any())
            .ToList();

        foreach (var start in starts)
        {
            var id = (string?)start.Attribute(W.id);
            if (id is null) continue;
            var end = ends.FirstOrDefault(e => (string?)e.Attribute(W.id) == id);
            var references = referenceRuns
                .Where(r => r.Descendants(W.commentReference)
                    .Any(cr => (string?)cr.Attribute(W.id) == id))
                .ToList();
            if (end is null || references.Count == 0) continue;

            start.Remove();
            end.Remove();
            foreach (var reference in references) reference.Remove();
            host.Add(start, end);
            host.Add(references);
        }
    }

    /// <summary>Whether the markup's payload survives resolution: an insertion survives
    /// accept, a deletion survives reject; move-source content is delete-like,
    /// move-destination content insert-like. The same matrix answers whether a revised
    /// paragraph mark keeps its paragraph alive.</summary>
    private static bool ContentSurvives(XName wrapperName, bool accept) =>
        wrapperName == W.ins || wrapperName == W.moveTo ? accept : !accept;

    private static bool Detached(XElement e) => e.Document is null;

    private static void UnwrapWrapper(XElement wrapper, bool restoreDeleted)
    {
        if (restoreDeleted)
        {
            foreach (var dt in wrapper.Descendants(W.delText).ToList())
                dt.ReplaceWith(new XElement(W.t, dt.Attributes(), dt.Nodes()));
            foreach (var di in wrapper.Descendants(W.delInstrText).ToList())
                di.ReplaceWith(new XElement(W.instrText, di.Attributes(), di.Nodes()));
        }
        // Detach the children before re-adding so they are MOVED, not cloned — a nested
        // revision element (a w:del inside an accepted w:ins) must stay the same live
        // node so a later resolution of ITS group still finds it attached.
        var nodes = wrapper.Nodes().ToList();
        wrapper.RemoveNodes();
        wrapper.ReplaceWith(nodes);
    }

    private static void RemoveContentWrapper(XElement wrapper)
    {
        var parent = wrapper.Parent;
        wrapper.Remove();
        // Collapse a hyperlink shell left with no content (mirrors RevisionProcessor's
        // wholly-deleted-hyperlink rule); bookmarks inside it are preserved.
        while (parent is not null && parent.Name == W.hyperlink
            && !parent.Elements().Any(e =>
                e.Name != W.bookmarkStart && e.Name != W.bookmarkEnd && e.Name != W.proofErr))
        {
            var grandParent = parent.Parent;
            var bookmarks = parent.Elements()
                .Where(e => e.Name == W.bookmarkStart || e.Name == W.bookmarkEnd)
                .ToList();
            foreach (var b in bookmarks) b.Remove();
            parent.ReplaceWith(bookmarks);
            parent = grandParent;
        }
    }

    private static void StripRowMark(XElement mark)
    {
        var trPr = mark.Parent;
        mark.Remove();
        if (trPr is not null && !trPr.HasElements && !trPr.HasAttributes) trPr.Remove();
    }

    private static void RemoveRow(XElement tr, List<XElement> removedBlocks)
    {
        var tbl = tr.Parent;
        tr.Remove();
        removedBlocks.Add(tr);
        // A table needs at least one row; resolving away the last one removes the table.
        if (tbl is not null && tbl.Name == W.tbl && !tbl.Elements(W.tr).Any())
        {
            tbl.Remove();
            removedBlocks.Add(tbl);
        }
    }

    /// <summary>Coalesce a paragraph whose mark is going away into the following paragraph
    /// (which keeps its own properties — the surviving-pilcrow rule RevisionProcessor and
    /// Word both follow). Returns the removed paragraph, or null when there is no
    /// following paragraph to merge into.</summary>
    private static XElement? MergeParagraphWithNext(XElement p)
    {
        XElement? next = null;
        foreach (var e in p.ElementsAfterSelf())
        {
            if (IgnorableBetween.Contains(e.Name)) continue;
            if (e.Name == W.p) next = e;
            break;
        }
        if (next is null) return null;

        var content = p.Elements().Where(e => e.Name != W.pPr).ToList();
        foreach (var c in content) c.Remove();
        var nextPPr = next.Element(W.pPr);
        if (nextPPr is not null) nextPPr.AddAfterSelf(content);
        else next.AddFirst(content);
        p.Remove();
        return p;
    }

    private static void CleanParagraphHusks(XElement p)
    {
        var pPr = p.Element(W.pPr);
        if (pPr is null) return;
        var rPr = pPr.Element(W.rPr);
        if (rPr is not null && !rPr.HasElements && !rPr.HasAttributes) rPr.Remove();
        if (!pPr.HasElements && !pPr.HasAttributes) pPr.Remove();
    }

    // ─── Format-change resolution ───────────────────────────────────────

    private static void AcceptProps(XElement change)
    {
        var parent = change.Parent;
        change.Remove();
        if (parent is not null && !parent.HasElements && !parent.HasAttributes
            && (parent.Name == W.rPr || parent.Name == W.pPr || parent.Name == W.trPr
                || parent.Name == W.tblPrEx))
        {
            parent.Remove();
        }
    }

    /// <summary>Restore the properties stored inside the <c>*PrChange</c> element, keeping
    /// the children the change's CT_*Base inner schema excludes (revision marks on a
    /// paragraph-mark rPr, header/footer references on sectPr, rPr/sectPr on pPr, row
    /// marks on trPr, cell revision marks on tcPr) in their schema position.</summary>
    private static void RejectProps(XElement change)
    {
        var parent = change.Parent;
        if (parent is null) { change.Remove(); return; }
        var stored = change.Elements().FirstOrDefault();
        var storedChildren = stored?.Elements().Select(e => new XElement(e)).ToList()
            ?? new List<XElement>();
        change.Remove();

        var pn = parent.Name;
        if (pn == W.rPr)
        {
            // CT_ParaRPr: ins/del/moveFrom/moveTo precede the property set.
            var marks = DetachAll(parent.Elements().Where(e => RevWrapperNames.Contains(e.Name)));
            parent.ReplaceNodes(marks, storedChildren);
        }
        else if (pn == W.pPr)
        {
            // CT_PPr: base props, then rPr, then sectPr (pPrChange's stored pPr is CT_PPrBase).
            var rPr = DetachAll(parent.Elements(W.rPr));
            var sectPr = DetachAll(parent.Elements(W.sectPr));
            parent.ReplaceNodes(storedChildren, rPr, sectPr);
        }
        else if (pn == W.sectPr)
        {
            // CT_SectPr: header/footer references come first (excluded from CT_SectPrBase).
            var refs = DetachAll(parent.Elements()
                .Where(e => e.Name == W.headerReference || e.Name == W.footerReference));
            parent.ReplaceNodes(refs, storedChildren);
        }
        else if (pn == W.trPr)
        {
            // CT_TrPr: base props, then ins/del row marks.
            var marks = DetachAll(parent.Elements().Where(e => e.Name == W.ins || e.Name == W.del));
            parent.ReplaceNodes(storedChildren, marks);
        }
        else if (pn == W.tcPr)
        {
            // CT_TcPr: base props, then cellIns/cellDel/cellMerge.
            var cellRevs = DetachAll(parent.Elements()
                .Where(e => e.Name == W.cellIns || e.Name == W.cellDel || e.Name == W.cellMerge));
            parent.ReplaceNodes(storedChildren, cellRevs);
        }
        else
        {
            // tblPr, tblGrid, tblPrEx — the stored element is the whole property set.
            parent.ReplaceNodes(storedChildren);
        }

        // A change whose stored old property set was empty leaves an empty husk —
        // remove it (Word writes no empty rPr/pPr), mirroring AcceptProps.
        if (!parent.HasElements && !parent.HasAttributes
            && (pn == W.rPr || pn == W.pPr || pn == W.trPr || pn == W.tblPrEx))
        {
            parent.Remove();
        }
    }

    private static List<XElement> DetachAll(IEnumerable<XElement> elements)
    {
        var list = elements.ToList();
        foreach (var e in list) e.Remove();
        return list;
    }
}
