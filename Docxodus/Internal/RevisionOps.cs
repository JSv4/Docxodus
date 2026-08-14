#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
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
/// Scope: run-content ins/del (any story), paragraph-mark ins/del, table-row
/// ins/del (<c>w:trPr</c> markers absorb their row's content markup), table-cell
/// insertion/deletion/vertical-merge operations, content-control envelope ranges,
/// numbering-property insertion/numbering cache changes, named move pairs (both
/// sides resolve together), and the format-change family
/// (<c>rPrChange</c>/<c>pPrChange</c>/<c>sectPrChange</c>/<c>tblPrChange</c>/
/// <c>trPrChange</c>/<c>tcPrChange</c>/<c>tblGridChange</c>/<c>tblPrExChange</c>).
/// Unsupported or malformed native markup is enumerated explicitly and fails closed.
/// </summary>
internal static class RevisionOps
{
    internal const string TypeInsert = "insert";
    internal const string TypeDelete = "delete";
    internal const string TypeMove = "move";
    internal const string TypeFormat = "format";
    internal const string TypeStructure = "structure";

    internal enum UnitKind
    {
        Content,
        ParaMark,
        RowMark,
        PropsChange,
        CellMark,
        NumberingPropertiesInsert,
        NumberingChange,
        StructuredRange,
        Unsupported,
    }

    /// <summary>One revision markup element, positioned in document order (a paragraph's
    /// mark unit is repositioned to the END of its paragraph — that is where the pilcrow
    /// lives semantically, and what makes multi-paragraph runs of markup group).</summary>
    internal sealed class RevisionUnit
    {
        required public XElement Element { get; init; }
        required public UnitKind Kind { get; init; }
        required public string Type { get; init; }
        required public RevisionFamily Family { get; init; }
        required public string Author { get; init; }
        public string? Date { get; init; }
        /// <summary>Move-range name when the unit sits inside a named move range — such
        /// units group per name (both sides of the pair) rather than by adjacency.</summary>
        public string? MoveName { get; init; }
        public XElement? Paragraph { get; init; }
        /// <summary>For RowMark: the <c>w:tr</c> itself. For Content: the marked row the
        /// unit sits inside (so the row group absorbs it), else null.</summary>
        public XElement? MarkedRow { get; init; }
        public XElement? MarkedCell { get; init; }
        public XElement? Table { get; init; }
        public XElement? StructuredWrapper { get; init; }
        public long? Wid { get; init; }
        public string? NativeId { get; init; }
    }

    internal sealed class RevisionGroup
    {
        public string Id { get; set; } = "";
        required public string Type { get; init; }
        required public RevisionFamily Family { get; init; }
        required public string Author { get; init; }
        public string? Date { get; set; }
        required public int PartIndex { get; init; }
        public string PartUri { get; set; } = "";
        public string Scope { get; set; } = "body";
        public RevisionResolutionStatus ResolutionStatus { get; set; } = RevisionResolutionStatus.Supported;
        public RevisionDiagnostic? Diagnostic { get; set; }
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

    private static readonly HashSet<XName> StructuredRangeNames = new()
    {
        W.customXmlDelRangeStart, W.customXmlDelRangeEnd,
        W.customXmlInsRangeStart, W.customXmlInsRangeEnd,
    };

    private static readonly HashSet<XName> UnsupportedRangeNames = new()
    {
        W.customXmlMoveFromRangeStart, W.customXmlMoveFromRangeEnd,
        W.customXmlMoveToRangeStart, W.customXmlMoveToRangeEnd,
    };

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

    internal static List<RevisionGroup> Enumerate(
        IReadOnlyList<(string PartUri, string Scope, XElement Root)> parts)
    {
        var groups = new List<RevisionGroup>();
        for (int pi = 0; pi < parts.Count; pi++)
        {
            var ctx = new WalkCtx();
            var units = new List<RevisionUnit>();
            WalkChildren(parts[pi].Root.Elements(), ctx, null, null, null, units);
            BuildGroups(units, ctx, pi, groups);
            AddStructuredRangeGroups(parts[pi].Root, pi, groups);
            AddUnsupportedGroups(parts[pi].Root, pi, groups);

            foreach (var group in groups.Where(g => g.PartIndex == pi))
            {
                group.PartUri = parts[pi].PartUri;
                group.Scope = parts[pi].Scope;
            }

            CoalesceTableStructureGroups(groups, pi);
            AbsorbTrackedStructuredPayload(groups, pi);
        }
        ValidateGroups(groups);
        AssignIds(groups);
        return groups;
    }

    private static void ValidateGroups(List<RevisionGroup> groups)
    {
        foreach (var group in groups.Where(g => g.ResolutionStatus == RevisionResolutionStatus.Supported))
        {
            if (group.Units.Any(u => string.IsNullOrEmpty(u.NativeId))
                && group.Units.Any(u => u.Kind != UnitKind.PropsChange
                    || u.Element.Name != W.tblGridChange))
            {
                group.ResolutionStatus = RevisionResolutionStatus.Malformed;
                group.Diagnostic = new RevisionDiagnostic(
                    "missing_revision_id",
                    "A live revision marker has no w:id and cannot be addressed stably.");
                continue;
            }

            if (group.Family == RevisionFamily.CellInsert
                || group.Family == RevisionFamily.CellDelete
                || group.Family == RevisionFamily.CellMerge)
            {
                var cells = group.Units.Where(u => u.Kind == UnitKind.CellMark).ToList();
                if (cells.Count == 0 || cells.Any(u => u.MarkedCell is null || u.Table is null))
                {
                    group.ResolutionStatus = RevisionResolutionStatus.Malformed;
                    group.Diagnostic = new RevisionDiagnostic(
                        "orphan_cell_revision",
                        "A cell structural marker is not a direct property of a table cell.");
                    continue;
                }

                if (group.Family == RevisionFamily.CellDelete)
                {
                    var deleted = cells.Select(u => u.MarkedCell!).ToHashSet();
                    bool invalidRow = deleted.GroupBy(c => c.Parent).Any(byRow =>
                    {
                        var rowCells = byRow.Key?.Elements(W.tc).ToList() ?? new List<XElement>();
                        int firstDeleted = rowCells.FindIndex(deleted.Contains);
                        return firstDeleted == 0 || rowCells.All(deleted.Contains);
                    });
                    if (invalidRow)
                    {
                        group.ResolutionStatus = RevisionResolutionStatus.Malformed;
                        group.Diagnostic = new RevisionDiagnostic(
                            "unabsorbable_cell_deletion",
                            "A deleted-cell run has no surviving predecessor that can absorb its grid columns.");
                        continue;
                    }
                }

                if (group.Family == RevisionFamily.CellMerge && cells.Any(u =>
                    (string?)u.Element.Attribute(W.vMerge) is not ("rest" or "cont")))
                {
                    group.ResolutionStatus = RevisionResolutionStatus.Malformed;
                    group.Diagnostic = new RevisionDiagnostic(
                        "invalid_cell_merge_state",
                        "w:cellMerge must carry w:vMerge='rest' or 'cont'.");
                }
            }

            if (group.Family == RevisionFamily.NumberingPropertiesInsert
                && group.Units.Any(u => u.Element.Parent?.Name != W.numPr
                    || u.Paragraph is null))
            {
                group.ResolutionStatus = RevisionResolutionStatus.Malformed;
                group.Diagnostic = new RevisionDiagnostic(
                    "orphan_numbering_revision",
                    "A numbering revision marker is not a direct child of paragraph w:numPr.");
            }

            if (group.Family == RevisionFamily.NumberingChange
                && group.Units.Any(u => (u.Element.Parent?.Name != W.numPr
                        && u.Element.Parent?.Name != W.fldChar)
                    || u.Paragraph is null))
            {
                group.ResolutionStatus = RevisionResolutionStatus.Malformed;
                group.Diagnostic = new RevisionDiagnostic(
                    "orphan_numbering_revision",
                    "w:numberingChange is not attached to paragraph numbering properties or a LISTNUM field.");
            }
        }

        // Reusing one live revision id for multiple independent groups in one part
        // makes an id-based operation inherently ambiguous. Range-pair duplication was
        // diagnosed earlier and groups already coalesced into one operation are fine.
        foreach (var collision in groups.SelectMany(g => ConstituentIds(g)
                .Select(id => (Group: g, Id: id)))
            .GroupBy(x => (x.Group.PartUri, x.Id))
            .Where(g => g.Select(x => x.Group).Distinct().Count() > 1))
        {
            foreach (var group in collision.Select(x => x.Group).Distinct()
                .Where(g => g.ResolutionStatus == RevisionResolutionStatus.Supported))
            {
                group.ResolutionStatus = RevisionResolutionStatus.Ambiguous;
                group.Diagnostic = new RevisionDiagnostic(
                    "duplicate_revision_id",
                    $"w:id '{collision.Key.Id}' identifies multiple live revisions in {collision.Key.PartUri}.");
            }
        }
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
            // Content-control envelope ranges are paired and validated in a dedicated
            // pass. Treating their starts as ordinary adjacent units loses the two-pair
            // topology that identifies the wrapper whose existence is revised.
            if (StructuredRangeNames.Contains(n) || UnsupportedRangeNames.Contains(n))
                continue;
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
            if ((n == W.cellIns || n == W.cellDel || n == W.cellMerge)
                && child.Parent?.Name == W.tcPr)
            {
                var cell = child.Ancestors(W.tc).FirstOrDefault();
                var table = child.Ancestors(W.tbl).FirstOrDefault();
                var family = n == W.cellIns ? RevisionFamily.CellInsert
                    : n == W.cellDel ? RevisionFamily.CellDelete
                    : RevisionFamily.CellMerge;
                sink.Add(new RevisionUnit
                {
                    Element = child,
                    Kind = UnitKind.CellMark,
                    Type = n == W.cellIns ? TypeInsert : n == W.cellDel ? TypeDelete : TypeStructure,
                    Family = family,
                    Author = AuthorOf(child),
                    Date = (string?)child.Attribute(W.date),
                    MarkedCell = cell,
                    MarkedRow = cell?.Parent,
                    Table = table,
                    Wid = WidOf(child),
                    NativeId = (string?)child.Attribute(W.id),
                });
                continue;
            }
            if (n == W.ins && child.Parent?.Name == W.numPr)
            {
                sink.Add(new RevisionUnit
                {
                    Element = child,
                    Kind = UnitKind.NumberingPropertiesInsert,
                    Type = TypeInsert,
                    Family = RevisionFamily.NumberingPropertiesInsert,
                    Author = AuthorOf(child),
                    Date = (string?)child.Attribute(W.date),
                    Paragraph = child.Ancestors(W.p).FirstOrDefault(),
                    Wid = WidOf(child),
                    NativeId = (string?)child.Attribute(W.id),
                });
                continue;
            }
            if (n == W.numberingChange)
            {
                sink.Add(new RevisionUnit
                {
                    Element = child,
                    Kind = UnitKind.NumberingChange,
                    Type = TypeFormat,
                    Family = RevisionFamily.NumberingChange,
                    Author = AuthorOf(child),
                    Date = (string?)child.Attribute(W.date),
                    Paragraph = child.Ancestors(W.p).FirstOrDefault(),
                    Wid = WidOf(child),
                    NativeId = (string?)child.Attribute(W.id),
                });
                continue;
            }
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

            // pPr is otherwise handled specially so paragraph-mark revisions can be
            // emitted at the semantic pilcrow position. Inventory the numbering-only
            // families explicitly, while excluding archived *PrChange payloads.
            foreach (var numPr in pPr.DescendantsAndSelf(W.numPr)
                .Where(np => !np.Ancestors().Any(a => PropsChangeNames.Contains(a.Name))))
            {
                foreach (var marker in numPr.Elements(W.ins))
                {
                    sink.Add(new RevisionUnit
                    {
                        Element = marker,
                        Kind = UnitKind.NumberingPropertiesInsert,
                        Type = TypeInsert,
                        Family = RevisionFamily.NumberingPropertiesInsert,
                        Author = AuthorOf(marker),
                        Date = (string?)marker.Attribute(W.date),
                        Paragraph = p,
                        Table = p.Ancestors(W.tbl).FirstOrDefault(),
                        Wid = WidOf(marker),
                        NativeId = (string?)marker.Attribute(W.id),
                    });
                }
                foreach (var marker in numPr.Elements(W.numberingChange))
                {
                    sink.Add(new RevisionUnit
                    {
                        Element = marker,
                        Kind = UnitKind.NumberingChange,
                        Type = TypeFormat,
                        Family = RevisionFamily.NumberingChange,
                        Author = AuthorOf(marker),
                        Date = (string?)marker.Attribute(W.date),
                        Paragraph = p,
                        Table = p.Ancestors(W.tbl).FirstOrDefault(),
                        Wid = WidOf(marker),
                        NativeId = (string?)marker.Attribute(W.id),
                    });
                }
            }
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
                    Family = rowType == TypeInsert ? RevisionFamily.RowInsert : RevisionFamily.RowDelete,
                    Author = AuthorOf(mark),
                    Date = (string?)mark.Attribute(W.date),
                    MarkedRow = tr,
                    Table = tr.Parent,
                    Wid = WidOf(mark),
                    NativeId = (string?)mark.Attribute(W.id),
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
        else if (kind == UnitKind.ParaMark && paragraph is not null &&
                 (n == W.moveFrom || n == W.moveTo))
        {
            // Whole-paragraph moves revise the paragraph mark in pPr/rPr. The named
            // range lives in the paragraph body and its stack has closed by the time
            // WalkParagraph emits this final pilcrow unit, so recover the name from
            // the sibling range start and keep the whole move one selectable revision.
            var startName = n == W.moveFrom ? W.moveFromRangeStart : W.moveToRangeStart;
            moveName = (string?)paragraph.Elements(startName).FirstOrDefault()?.Attribute(W.name);
        }
        return new RevisionUnit
        {
            Element = el,
            Kind = kind,
            Type = type,
            Family = kind == UnitKind.ParaMark
                ? RevisionFamily.ParagraphMark
                : type == TypeInsert ? RevisionFamily.ContentInsert
                : type == TypeDelete ? RevisionFamily.ContentDelete
                : RevisionFamily.Move,
            Author = AuthorOf(el),
            Date = (string?)el.Attribute(W.date),
            MoveName = moveName,
            Paragraph = paragraph,
            MarkedRow = markedRowType == type ? markedRow : null,
            Table = el.Ancestors(W.tbl).FirstOrDefault(),
            Wid = WidOf(el),
            NativeId = (string?)el.Attribute(W.id),
        };
    }

    private static RevisionUnit MakePropsUnit(XElement el, XElement? paragraph) =>
        new()
        {
            Element = el,
            Kind = UnitKind.PropsChange,
            Type = TypeFormat,
            Family = RevisionFamily.PropertiesChange,
            Author = AuthorOf(el),
            Date = (string?)el.Attribute(W.date),
            Paragraph = paragraph,
            Table = el.Ancestors(W.tbl).FirstOrDefault(),
            Wid = WidOf(el),
            NativeId = (string?)el.Attribute(W.id),
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
        var cellGroups = new List<RevisionGroup>();
        RevisionGroup? cur = null;
        RevisionGroup? lastRowGroup = null;

        foreach (var u in units)
        {
            if (u.MoveName is not null)
            {
                if (!moveGroups.TryGetValue(u.MoveName, out var mg))
                {
                    mg = new RevisionGroup
                    {
                        Type = TypeMove,
                        Family = RevisionFamily.Move,
                        Author = u.Author,
                        Date = u.Date,
                        PartIndex = partIndex,
                    };
                    if (ctx.RangeMarkers.TryGetValue(u.MoveName, out var markers))
                        mg.RangeMarkers.AddRange(markers);
                    moveGroups[u.MoveName] = mg;
                    groups.Add(mg);
                }
                mg.Units.Add(u);
                continue;
            }

            if (u.Kind == UnitKind.CellMark)
            {
                var cellGroup = cellGroups.FirstOrDefault(g =>
                    g.Family == u.Family && g.Author == u.Author && g.Date == u.Date
                    && ReferenceEquals(g.Units[0].Table, u.Table));
                if (cellGroup is null)
                {
                    cellGroup = NewGroup(u, partIndex);
                    cellGroups.Add(cellGroup);
                    groups.Add(cellGroup);
                }
                else
                {
                    cellGroup.Units.Add(u);
                }
                cur = null;
                continue;
            }

            if (u.Kind == UnitKind.NumberingPropertiesInsert || u.Kind == UnitKind.NumberingChange)
            {
                groups.Add(NewGroup(u, partIndex));
                cur = null;
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
        var g = new RevisionGroup
        {
            Type = u.Type,
            Family = u.Family,
            Author = u.Author,
            Date = u.Date,
            PartIndex = partIndex,
        };
        g.Units.Add(u);
        return g;
    }

    /// <summary>
    /// Recognize the exact two-range topology Word uses to revise an SDT envelope.
    /// Any unpaired, duplicated, or topologically misplaced marker remains visible as
    /// a malformed/ambiguous registry entry instead of being silently ignored.
    /// </summary>
    private static void AddStructuredRangeGroups(
        XElement root, int partIndex, List<RevisionGroup> groups)
    {
        var allMarkers = root.Descendants()
            .Where(e => StructuredRangeNames.Contains(e.Name))
            .ToList();
        var used = new HashSet<XElement>();

        foreach (var sdt in root.Descendants(W.sdt))
        {
            var content = sdt.Element(W.sdtContent);
            if (content is null) continue;

            var before = sdt.ElementsBeforeSelf().LastOrDefault();
            var after = sdt.ElementsAfterSelf().FirstOrDefault();
            var firstInside = content.Elements().FirstOrDefault();
            var lastInside = content.Elements().LastOrDefault();
            if (before is null || after is null || firstInside is null || lastInside is null)
                continue;

            bool isInsert = before.Name == W.customXmlInsRangeStart;
            bool isDelete = before.Name == W.customXmlDelRangeStart;
            if (!isInsert && !isDelete) continue;

            var startName = isInsert ? W.customXmlInsRangeStart : W.customXmlDelRangeStart;
            var endName = isInsert ? W.customXmlInsRangeEnd : W.customXmlDelRangeEnd;
            var firstId = (string?)before.Attribute(W.id);
            var secondId = (string?)lastInside.Attribute(W.id);
            if (firstInside.Name != endName || lastInside.Name != startName || after.Name != endName
                || string.IsNullOrEmpty(firstId) || string.IsNullOrEmpty(secondId)
                || (string?)firstInside.Attribute(W.id) != firstId
                || (string?)after.Attribute(W.id) != secondId)
            {
                continue;
            }

            // Both starts describe one wrapper revision and must carry a coherent stamp.
            var author = AuthorOf(before);
            var date = (string?)before.Attribute(W.date);
            if (AuthorOf(lastInside) != author || (string?)lastInside.Attribute(W.date) != date)
                continue;

            var family = isInsert
                ? RevisionFamily.ContentControlInsert
                : RevisionFamily.ContentControlDelete;
            var unit = new RevisionUnit
            {
                Element = before,
                Kind = UnitKind.StructuredRange,
                Type = isInsert ? TypeInsert : TypeDelete,
                Family = family,
                Author = author,
                Date = date,
                Paragraph = sdt.AncestorsAndSelf(W.p).FirstOrDefault(),
                MarkedCell = sdt.Ancestors(W.tc).FirstOrDefault(),
                MarkedRow = sdt.Ancestors(W.tr).FirstOrDefault(),
                Table = sdt.Ancestors(W.tbl).FirstOrDefault(),
                StructuredWrapper = sdt,
                Wid = WidOf(before),
                NativeId = firstId,
            };
            var group = NewGroup(unit, partIndex);
            group.RangeMarkers.AddRange(new[] { before, firstInside, lastInside, after });
            groups.Add(group);
            used.UnionWith(group.RangeMarkers);
        }

        foreach (var markerGroup in allMarkers.Where(m => !used.Contains(m))
            .GroupBy(m => (Family: RangeFamily(m.Name), Id: (string?)m.Attribute(W.id))))
        {
            var markers = markerGroup.ToList();
            var starts = markers.Where(m => m.Name.LocalName.EndsWith("RangeStart", StringComparison.Ordinal)).ToList();
            var family = markerGroup.Key.Family;
            var exemplar = starts.FirstOrDefault() ?? markers[0];
            var unit = new RevisionUnit
            {
                Element = exemplar,
                Kind = UnitKind.StructuredRange,
                Type = family == RevisionFamily.ContentControlInsert ? TypeInsert : TypeDelete,
                Family = family,
                Author = AuthorOf(exemplar),
                Date = (string?)exemplar.Attribute(W.date),
                Paragraph = exemplar.Ancestors(W.p).FirstOrDefault(),
                MarkedCell = exemplar.Ancestors(W.tc).FirstOrDefault(),
                MarkedRow = exemplar.Ancestors(W.tr).FirstOrDefault(),
                Table = exemplar.Ancestors(W.tbl).FirstOrDefault(),
                Wid = WidOf(exemplar),
                NativeId = (string?)exemplar.Attribute(W.id),
            };
            var group = NewGroup(unit, partIndex);
            group.RangeMarkers.AddRange(markers);
            bool duplicate = markers.Count(m => m.Name.LocalName.EndsWith("RangeStart", StringComparison.Ordinal)) > 1
                || markers.Count(m => m.Name.LocalName.EndsWith("RangeEnd", StringComparison.Ordinal)) > 1;
            group.ResolutionStatus = duplicate
                ? RevisionResolutionStatus.Ambiguous
                : RevisionResolutionStatus.Malformed;
            group.Diagnostic = new RevisionDiagnostic(
                duplicate ? "duplicate_range_id" : "malformed_range_topology",
                duplicate
                    ? "Content-control revision range id is duplicated in its owning part."
                    : "Content-control revision ranges do not form Word's exact two-pair SDT envelope topology.");
            groups.Add(group);
        }
    }

    private static RevisionFamily RangeFamily(XName name) =>
        name == W.customXmlInsRangeStart || name == W.customXmlInsRangeEnd
            ? RevisionFamily.ContentControlInsert
            : RevisionFamily.ContentControlDelete;

    private static void AddUnsupportedGroups(XElement root, int partIndex, List<RevisionGroup> groups)
    {
        foreach (var markerGroup in root.Descendants()
            .Where(e => UnsupportedRangeNames.Contains(e.Name))
            .GroupBy(e => ((string?)e.Attribute(W.id), e.Name.LocalName.Contains("MoveFrom", StringComparison.Ordinal))))
        {
            var markers = markerGroup.ToList();
            var exemplar = markers.FirstOrDefault(m => m.Name.LocalName.EndsWith("RangeStart", StringComparison.Ordinal))
                ?? markers[0];
            var unit = new RevisionUnit
            {
                Element = exemplar,
                Kind = UnitKind.Unsupported,
                Type = TypeMove,
                Family = RevisionFamily.Unsupported,
                Author = AuthorOf(exemplar),
                Date = (string?)exemplar.Attribute(W.date),
                Paragraph = exemplar.Ancestors(W.p).FirstOrDefault(),
                Table = exemplar.Ancestors(W.tbl).FirstOrDefault(),
                Wid = WidOf(exemplar),
                NativeId = (string?)exemplar.Attribute(W.id),
            };
            var group = NewGroup(unit, partIndex);
            group.RangeMarkers.AddRange(markers);
            group.ResolutionStatus = RevisionResolutionStatus.Unsupported;
            group.Diagnostic = new RevisionDiagnostic(
                "unsupported_custom_xml_move_range",
                "customXml move-range revisions are listed but cannot be selectively resolved.");
            groups.Add(group);
        }
    }

    /// <summary>
    /// Word records one cell-structure action as live cell marks plus associated table,
    /// cell, paragraph-property, and content revisions. Fold that coherent stamp into
    /// one atomic registry entry. Archived markers inside *PrChange payloads were never
    /// walked, so they cannot be mistaken for live operations.
    /// </summary>
    private static void CoalesceTableStructureGroups(List<RevisionGroup> groups, int partIndex)
    {
        var cellGroups = groups.Where(g => g.PartIndex == partIndex
            && (g.Family == RevisionFamily.CellInsert
                || g.Family == RevisionFamily.CellDelete
                || g.Family == RevisionFamily.CellMerge))
            .ToList();

        foreach (var byTable in cellGroups.GroupBy(g => g.Units[0].Table))
        {
            if (byTable.Key is null) continue;
            var tableCellGroups = byTable.ToList();
            var candidates = groups.Where(g => g.PartIndex == partIndex
                && !tableCellGroups.Contains(g)
                && g.Units.Count > 0
                && g.Units.All(u => ReferenceEquals(u.Table, byTable.Key)))
                .ToList();

            foreach (var candidate in candidates)
            {
                var matches = tableCellGroups.Where(c =>
                    c.Author == candidate.Author && c.Date == candidate.Date).ToList();
                if (matches.Count == 0 && candidate.Units.All(u => u.Element.Name == W.tblGridChange)
                    && candidate.Author == "unknown" && candidate.Date is null)
                {
                    matches = tableCellGroups;
                }

                if (matches.Count == 1)
                {
                    matches[0].Units.AddRange(candidate.Units);
                    matches[0].RangeMarkers.AddRange(candidate.RangeMarkers);
                    groups.Remove(candidate);
                }
                else if (matches.Count > 1)
                {
                    candidate.ResolutionStatus = RevisionResolutionStatus.Ambiguous;
                    candidate.Diagnostic = new RevisionDiagnostic(
                        "ambiguous_table_structure_cluster",
                        "An unattributed table property revision matches multiple live cell operations.");
                    foreach (var match in matches)
                    {
                        match.ResolutionStatus = RevisionResolutionStatus.Ambiguous;
                        match.Diagnostic = candidate.Diagnostic;
                    }
                }
            }
        }
    }

    private static void AbsorbTrackedStructuredPayload(List<RevisionGroup> groups, int partIndex)
    {
        foreach (var structured in groups.Where(g => g.PartIndex == partIndex
            && (g.Family == RevisionFamily.ContentControlInsert
                || g.Family == RevisionFamily.ContentControlDelete)
            && g.ResolutionStatus == RevisionResolutionStatus.Supported).ToList())
        {
            var wrapper = structured.Units[0].StructuredWrapper;
            if (wrapper is null) continue;
            var candidates = groups.Where(g => g != structured && g.PartIndex == partIndex
                && g.Type == structured.Type && g.Author == structured.Author && g.Date == structured.Date
                && g.Units.Count > 0
                && g.Units.All(u => u.Element.AncestorsAndSelf().Contains(wrapper)))
                .ToList();
            foreach (var candidate in candidates)
            {
                structured.Units.AddRange(candidate.Units);
                structured.RangeMarkers.AddRange(candidate.RangeMarkers);
                groups.Remove(candidate);
            }
        }
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
        groups.Sort((a, b) =>
        {
            int part = a.PartIndex.CompareTo(b.PartIndex);
            if (part != 0) return part;
            var ae = a.Units.FirstOrDefault()?.Element ?? a.RangeMarkers.FirstOrDefault();
            var be = b.Units.FirstOrDefault()?.Element ?? b.RangeMarkers.FirstOrDefault();
            if (ae is null || be is null) return ae is null ? (be is null ? 0 : 1) : -1;
            return XNode.DocumentOrderComparer.Compare(ae, be);
        });

        var identityMaterial = groups.ToDictionary(g => g, g =>
        {
            var constituents = ConstituentKeys(g);
            return g.PartUri + "\n" + g.Family + "\n" + string.Join("\n", constituents);
        });

        foreach (var candidate in groups.GroupBy(g => StableId(identityMaterial[g])))
        {
            if (candidate.Count() == 1)
            {
                candidate.First().Id = candidate.Key;
                continue;
            }

            // Invalid documents can reuse one native id for independent live operations.
            // They remain fail-closed/ambiguous, but list ids must still be unique so a
            // transport cannot silently overwrite one entry in an id-keyed map. The
            // collision ordinal is stable under resolution of unrelated revisions.
            int ordinal = 0;
            foreach (var group in candidate)
                group.Id = StableId(identityMaterial[group] + "\ncollision:" + ordinal++);
        }
    }

    private static string StableId(string material)
    {
        var digest = SHA256.HashData(Encoding.UTF8.GetBytes(material));
        // Opaque, part-qualified, deterministic identity. Twenty hex characters
        // provide 80 bits while keeping transport payloads compact.
        return "rev2-" + Convert.ToHexStringLower(digest.AsSpan(0, 10));
    }

    internal static IReadOnlyList<string> ConstituentIds(RevisionGroup group) =>
        group.Units.Select(u => u.NativeId)
            .Concat(group.RangeMarkers.Select(m => (string?)m.Attribute(W.id)))
            .Where(id => !string.IsNullOrEmpty(id))
            .Select(id => id!)
            .Distinct(StringComparer.Ordinal)
            .OrderBy(id => long.TryParse(id, out var n) ? n : long.MaxValue)
            .ThenBy(id => id, StringComparer.Ordinal)
            .ToList();

    internal static string? LegacyId(RevisionGroup group)
    {
        var ids = ConstituentIds(group);
        var numeric = ids.Select(id => long.TryParse(id, out var value) ? value : (long?)null)
            .Where(value => value.HasValue)
            .Select(value => value!.Value)
            .DefaultIfEmpty()
            .Min();
        return ids.Any(id => long.TryParse(id, out _))
            ? "rev" + numeric.ToString(System.Globalization.CultureInfo.InvariantCulture)
            : null;
    }

    private static IReadOnlyList<string> ConstituentKeys(RevisionGroup group)
    {
        var keys = group.Units.Select(u =>
                u.Element.Name.NamespaceName + ":" + u.Element.Name.LocalName + ":"
                + (u.NativeId ?? ElementPath(u.Element)))
            .Concat(group.RangeMarkers.Select(m =>
                m.Name.NamespaceName + ":" + m.Name.LocalName + ":"
                + ((string?)m.Attribute(W.id) ?? ElementPath(m))))
            .Distinct(StringComparer.Ordinal)
            .OrderBy(k => k, StringComparer.Ordinal)
            .ToList();
        return keys.Count > 0 ? keys : new[] { "empty" };
    }

    private static string ElementPath(XElement element)
    {
        var segments = new Stack<string>();
        for (var current = element; current is not null; current = current.Parent)
        {
            int index = current.ElementsBeforeSelf(current.Name).Count();
            segments.Push(current.Name.LocalName + "[" + index + "]");
        }
        return string.Join("/", segments);
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
                case UnitKind.CellMark:
                    if (u.MarkedCell is { } cell)
                        AppendVisibleText(cell, u.Family == RevisionFamily.CellDelete ? W.delText : W.t, sb);
                    break;
                case UnitKind.NumberingPropertiesInsert:
                case UnitKind.NumberingChange:
                    if (u.Paragraph is { } numberedParagraph)
                        AppendVisibleText(numberedParagraph, W.t, sb);
                    break;
                case UnitKind.StructuredRange:
                    if (u.StructuredWrapper is { } wrapper)
                        AppendVisibleText(wrapper, W.t, sb);
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
        if (g.ResolutionStatus != RevisionResolutionStatus.Supported)
            throw new InvalidOperationException(g.Diagnostic?.Message
                ?? "revision cannot be resolved safely");

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

        foreach (var u in g.Units.Where(u => u.Kind == UnitKind.NumberingPropertiesInsert))
        {
            if (Detached(u.Element)) continue;
            var numPr = u.Element.Parent;
            if (accept)
                u.Element.Remove();
            else if (numPr?.Name == W.numPr)
                numPr.Remove();
        }

        foreach (var u in g.Units.Where(u => u.Kind == UnitKind.NumberingChange))
            if (!Detached(u.Element)) u.Element.Remove();

        ResolveCellStructure(g, accept, removedBlocks);

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

        ResolveStructuredWrapper(g, accept, removedBlocks);

        foreach (var m in g.RangeMarkers)
            if (!Detached(m)) m.Remove();

        foreach (var p in touchedParagraphs)
        {
            if (Detached(p)) continue;
            CleanParagraphHusks(p);
        }

        return removedBlocks;
    }

    private static void ResolveCellStructure(
        RevisionGroup group, bool accept, List<XElement> removedElements)
    {
        var cellUnits = group.Units.Where(u => u.Kind == UnitKind.CellMark).ToList();
        if (cellUnits.Count == 0) return;

        if (group.Family == RevisionFamily.CellInsert)
        {
            if (accept)
            {
                foreach (var unit in cellUnits)
                    if (!Detached(unit.Element)) unit.Element.Remove();
            }
            else
            {
                foreach (var cell in cellUnits.Select(u => u.MarkedCell)
                    .Where(c => c is not null).Select(c => c!).Distinct().ToList())
                {
                    if (Detached(cell)) continue;
                    cell.Remove();
                    removedElements.Add(cell);
                }
            }
        }
        else if (group.Family == RevisionFamily.CellDelete)
        {
            if (accept)
                AcceptDeletedCells(cellUnits, removedElements);
            else
                foreach (var unit in cellUnits)
                    if (!Detached(unit.Element)) unit.Element.Remove();
        }
        else if (group.Family == RevisionFamily.CellMerge)
        {
            foreach (var unit in cellUnits)
            {
                if (Detached(unit.Element)) continue;
                if (accept)
                {
                    var revised = (string?)unit.Element.Attribute(W.vMerge);
                    if (revised == "rest")
                        unit.Element.ReplaceWith(new XElement(W.vMerge,
                            new XAttribute(W.val, "restart")));
                    else if (revised == "cont")
                        unit.Element.ReplaceWith(new XElement(W.vMerge,
                            new XAttribute(W.val, "continue")));
                    else
                        unit.Element.Remove();
                }
                else
                {
                    var original = (string?)unit.Element.Attribute(W.vMergeOrig);
                    if (original == "rest")
                        unit.Element.ReplaceWith(new XElement(W.vMerge,
                            new XAttribute(W.val, "restart")));
                    else if (original == "cont")
                        unit.Element.ReplaceWith(new XElement(W.vMerge,
                            new XAttribute(W.val, "continue")));
                    else
                        unit.Element.Remove();
                }
            }
        }

        // Rejecting associated tcPrChange revisions can expose archived structural
        // marks from the old property shell. They belong to the same operation but are
        // not live revisions; remove only marks carrying this exact operation stamp.
        if (!accept)
        {
            var structuralName = group.Family == RevisionFamily.CellInsert ? W.cellIns
                : group.Family == RevisionFamily.CellDelete ? W.cellDel
                : W.cellMerge;
            foreach (var table in cellUnits.Select(u => u.Table).Where(t => t is not null)
                .Select(t => t!).Distinct())
            {
                foreach (var marker in table.Descendants()
                    .Where(e => e.Name == structuralName)
                    .Where(e => AuthorOf(e) == group.Author
                        && (string?)e.Attribute(W.date) == group.Date).ToList())
                    marker.Remove();
            }
        }
    }

    /// <summary>
    /// Accept cell deletion using grid units, not physical-cell count. Consecutive
    /// deleted cells contribute the sum of their pre-existing gridSpan values to the
    /// nearest surviving predecessor.
    /// </summary>
    private static void AcceptDeletedCells(
        IReadOnlyList<RevisionUnit> units, List<XElement> removedElements)
    {
        var deleted = units.Select(u => u.MarkedCell).Where(c => c is not null)
            .Select(c => c!).ToHashSet();
        foreach (var row in deleted.Select(c => c.Parent).Where(r => r is not null)
            .Select(r => r!).Distinct().ToList())
        {
            var cells = row.Elements(W.tc).ToList();
            XElement? predecessor = null;
            int pendingSpan = 0;
            foreach (var cell in cells)
            {
                if (deleted.Contains(cell))
                {
                    pendingSpan += CellGridSpan(cell);
                    cell.Remove();
                    removedElements.Add(cell);
                    continue;
                }

                if (pendingSpan > 0)
                {
                    if (predecessor is null)
                        throw new InvalidOperationException(
                            "A deleted-cell run has no surviving predecessor to absorb its grid columns.");
                    SetCellGridSpan(predecessor, CellGridSpan(predecessor) + pendingSpan);
                    pendingSpan = 0;
                }
                predecessor = cell;
            }
            if (pendingSpan > 0)
            {
                if (predecessor is null)
                    throw new InvalidOperationException(
                        "Resolving the cell deletion would remove every cell in a row.");
                SetCellGridSpan(predecessor, CellGridSpan(predecessor) + pendingSpan);
            }
        }
    }

    private static int CellGridSpan(XElement cell) =>
        Math.Max(1, (int?)cell.Element(W.tcPr)?.Element(W.gridSpan)?.Attribute(W.val) ?? 1);

    private static void SetCellGridSpan(XElement cell, int span)
    {
        var tcPr = cell.Element(W.tcPr);
        if (tcPr is null)
        {
            tcPr = new XElement(W.tcPr);
            cell.AddFirst(tcPr);
        }
        var gridSpan = tcPr.Element(W.gridSpan);
        if (span <= 1)
        {
            gridSpan?.Remove();
            return;
        }
        if (gridSpan is null)
        {
            gridSpan = new XElement(W.gridSpan, new XAttribute(W.val, span));
            var before = tcPr.Elements().FirstOrDefault(e =>
                e.Name == W.hMerge || e.Name == W.vMerge || e.Name == W.tcBorders
                || e.Name == W.shd || e.Name == W.noWrap || e.Name == W.tcMar
                || e.Name == W.textDirection || e.Name == W.tcFitText
                || e.Name == W.vAlign || e.Name == W.hideMark
                || e.Name == W.cellIns || e.Name == W.cellDel || e.Name == W.cellMerge
                || e.Name == W.tcPrChange);
            if (before is null) tcPr.Add(gridSpan);
            else before.AddBeforeSelf(gridSpan);
        }
        else
        {
            gridSpan.SetAttributeValue(W.val, span);
        }
    }

    private static void ResolveStructuredWrapper(
        RevisionGroup group, bool accept, List<XElement> removedElements)
    {
        if (group.Family != RevisionFamily.ContentControlInsert
            && group.Family != RevisionFamily.ContentControlDelete)
            return;

        var wrapper = group.Units.FirstOrDefault(u => u.Kind == UnitKind.StructuredRange)
            ?.StructuredWrapper;
        if (wrapper is null || Detached(wrapper)) return;

        bool wrapperSurvives = group.Family == RevisionFamily.ContentControlInsert
            ? accept
            : !accept;
        if (wrapperSurvives) return;

        var content = wrapper.Element(W.sdtContent);
        var nodes = content?.Nodes().Where(n => n is not XElement e
            || !StructuredRangeNames.Contains(e.Name)).ToList() ?? new List<XNode>();
        foreach (var node in nodes) node.Remove();
        wrapper.ReplaceWith(nodes);
        removedElements.Add(wrapper);
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
        if (rPr is not null && !rPr.HasElements && HasNoSemanticAttributes(rPr)) rPr.Remove();
        if (!pPr.HasElements && HasNoSemanticAttributes(pPr)) pPr.Remove();
    }

    private static bool HasNoSemanticAttributes(XElement element) =>
        element.Attributes().All(a => a.IsNamespaceDeclaration || a.Name == PtOpenXml.Unid);

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
