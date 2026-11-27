// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Buffers.Binary;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Docxodus.Internal;
using Docxodus.Verification;
using DocumentFormat.OpenXml.Packaging;
using GridCell = Docxodus.Internal.TableGridCell;

namespace Docxodus;

public sealed partial class DocxSession
{
    // ─── Tier B: structural ops ──────────────────────────────────────────

    /// <summary>
    /// Reorder one top-level body block relative to another. Paragraphs, headings,
    /// list items, and whole tables are supported. A direct edit moves the existing
    /// XML element; <see cref="TrackedChangeMode.RenderInline"/> emits a native named
    /// paragraph move or a Word-native deleted/inserted table pair. What a bookmark, comment,
    /// permission or field covers never changes: a marker between blocks goes with the block when
    /// its other end does, and any other such move is refused. Where a move would leave two tables
    /// touching, or a table ending the body, the empty paragraph it adds is in
    /// <see cref="EditResult.Created"/> (after the destination of a tracked move).
    /// </summary>
    public EditResult MoveBlock(string sourceAnchorId, string targetAnchorId, Position pos)
    {
        if (MutationRefusal() is { } refusal) return EditResult.Fail(refusal);

        var sourceTarget = FindAnchor(sourceAnchorId);
        if (sourceTarget is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound,
                $"source anchor not found: {sourceAnchorId}", sourceAnchorId);
        var targetTarget = FindAnchor(targetAnchorId);
        if (targetTarget is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound,
                $"target anchor not found: {targetAnchorId}", targetAnchorId);

        if (sourceTarget.Anchor.Kind is not ("p" or "h" or "li" or "tbl"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"MoveBlock requires a paragraph/heading/list/table source; got kind={sourceTarget.Anchor.Kind}",
                sourceAnchorId);
        if (targetTarget.Anchor.Kind is not ("p" or "h" or "li" or "tbl"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"MoveBlock requires a paragraph/heading/list/table target; got kind={targetTarget.Anchor.Kind}",
                targetAnchorId);
        if (sourceTarget.Anchor.Scope != targetTarget.Anchor.Scope)
            return EditResult.Fail(EditErrorCode.InvalidPosition,
                "MoveBlock source and target must be in the same package part", sourceAnchorId);

        var source = sourceTarget.Resolve(_doc!);
        var target = targetTarget.Resolve(_doc!);
        if (source is null || target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "source or target element resolved null", sourceAnchorId);
        if (ReferenceEquals(source, target))
            return new EditResult { Success = true };
        if (!ReferenceEquals(source.Parent, target.Parent) || source.Parent is not { } parent)
            return EditResult.Fail(EditErrorCode.InvalidPosition,
                "MoveBlock source and target must share a direct XML parent", sourceAnchorId);
        if (!IsEditorBodyBlock(source, parent) || !IsEditorBodyBlock(target, parent))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                "MoveBlock only supports top-level body blocks (including flattened body content controls)",
                sourceAnchorId);

        var context = BuildBlockMoveContext(parent);
        var run = MovingRun(context, context.Index[source], context.Index[source]);
        var landing = LandRun(context, run, context.Index[target], pos);
        // A source already in the requested slot is a true no-op: do not consume undo.
        if (landing.InPlace) return new EditResult { Success = true };

        if (TrackedBookmarkMoveRejection(source) is { } bookmarkRejection)
            return EditResult.Fail(EditErrorCode.UnsupportedInlineBoundary, bookmarkRejection, sourceAnchorId);

        if (MoveSourceRejection(context, run) is { } sourceRejection)
            return EditResult.Fail(EditErrorCode.InvalidPosition, sourceRejection, sourceAnchorId);
        if (landing.Error is { } safetyError)
            return EditResult.Fail(EditErrorCode.InvalidPosition, safetyError, sourceAnchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            if (_trackedChanges != TrackedChangeMode.RenderInline)
            {
                var added = RelocateRun(context, run, landing.Slot);
                InvalidateProjectionCache();
                return new EditResult
                {
                    Success = true,
                    Modified = new[] { sourceTarget.Anchor },
                    Created = added.Count == 0 ? Array.Empty<Anchor>() : ResolveAnchorsForElements(added),
                    Patch = PatchFor(sourceTarget),
                };
            }

            var destination = new XElement(source);
            foreach (var el in destination.DescendantsAndSelf())
                el.Attributes(PtOpenXml.Unid).Remove();
            Internal.RevisionOps.RenumberCopiedRevisions(destination, NextRevisionId);

            // Both copies are live while the revision is pending, so the shared comment ids have to
            // be split. The move SOURCE takes the clones (see CloneCommentsForMoveSource), leaving
            // the destination wired to the original comments and their threads.
            if (_doc!.MainDocumentPart is { } commentHost)
                Internal.CommentOps.CloneCommentsForMoveSource(commentHost, source, NextRevisionId);

            var author = _revisionAuthor ?? "docxodus";
            var date = NextTrackedFormatRevisionDate();
            if (source.Name == W.p)
            {
                var moveName = $"move{NextRevisionId()}";
                MarkParagraphAsTrackedMove(source, from: true, moveName, author, date);
                MarkParagraphAsTrackedMove(destination, from: false, moveName, author, date);
            }
            else
            {
                MarkTableRowsAsTrackedRevision(source, inserted: false, author, date);
                MarkTableRowsAsTrackedRevision(destination, inserted: true, author, date);
            }

            UnidHelper.AssignToSelfAndDescendants(destination);
            InsertAtSlot(context, run, landing.Slot, destination);
            // The paragraphs that keep tables apart once the move is accepted are inserted with it.
            var leftBehind = new XElement(W.p);
            source.AddAfterSelf(leftBehind);
            var separators = KeepTablesApart(leftBehind, destination, destination);
            foreach (var separator in separators) MarkParagraphMark(separator, W.ins, author, date);

            var destinationUnid = (string)destination.Attribute(PtOpenXml.Unid)!;
            var destinationAnchor = new Anchor(
                $"{sourceTarget.Anchor.Kind}:{sourceTarget.Anchor.Scope}:{destinationUnid}",
                sourceTarget.Anchor.Kind, sourceTarget.Anchor.Scope, destinationUnid);

            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Modified = new[] { sourceTarget.Anchor },
                Created = separators.Count == 0 ? new[] { destinationAnchor }
                    : ResolveAnchorsForElements(separators).Prepend(destinationAnchor).ToArray(),
                Patch = PatchFor(sourceTarget),
            };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, sourceAnchorId);
        }
    }

    /// <summary>A native tracked move keeps source and destination copies live simultaneously.
    /// Duplicating bookmark names violates global bookmark identity; moving the markers to only one
    /// side would lose them on either accept or reject. Reject explicitly instead of emitting an
    /// ambiguous pending document.</summary>
    private string? TrackedBookmarkMoveRejection(XElement source) =>
        _trackedChanges == TrackedChangeMode.RenderInline
        && source.DescendantsAndSelf().Any(e => e.Name == W.bookmarkStart || e.Name == W.bookmarkEnd)
            ? "tracked block moves containing bookmark markers are unsupported because both revision sides are live"
            : null;

    /// <summary>Reject reasons that depend only on the SOURCE block — if one applies, the block
    /// cannot be moved anywhere, so <see cref="ValidMoveTargets"/> can answer with an empty set
    /// without testing a single target. Shared with <see cref="MoveBlock"/> so the drag UI and the
    /// engine can never disagree about what is movable.</summary>
    private string? MoveSourceRejection(XElement source)
    {
        if (source.Name == W.p && source.Element(W.pPr)?.Element(W.sectPr) is not null)
            return "cannot move a paragraph that owns a section break";

        // MoveBlock tests this first and reports it as UnsupportedInlineBoundary; repeating the
        // PREDICATE (not the check) here is what keeps ValidMoveTargets from advertising a drop the
        // engine will refuse. One owner, two callers, two error codes.
        if (TrackedBookmarkMoveRejection(source) is { } bookmarkRejection)
            return bookmarkRejection;

        // Re-wrapping an existing revision can create illegal nested move markup. Direct
        // mode can carry ordinary ins/del content safely, but an existing named move range
        // is tied to its current document-order location and is intentionally immovable.
        if (source.DescendantsAndSelf().Any(e =>
                e.Name == W.moveFromRangeStart || e.Name == W.moveFromRangeEnd ||
                e.Name == W.moveToRangeStart || e.Name == W.moveToRangeEnd))
            return "cannot move a block that is already part of a native move range";
        if (_trackedChanges == TrackedChangeMode.RenderInline &&
            source.DescendantsAndSelf().Any(e =>
                e.Name == W.ins || e.Name == W.del || e.Name == W.moveFrom || e.Name == W.moveTo))
            return "cannot create a tracked move around a block that already contains revisions";
        return null;
    }

    private string? MoveSourceRejection(BlockMoveContext context, (int First, int Last) run) =>
        context.Run(run).Select(MoveSourceRejection).FirstOrDefault(why => why is not null);

    /// <summary>
    /// The anchors this block may legally be moved next to, in document order — the drop targets a
    /// drag UI should offer. Empty when the block cannot move at all (it owns a section break, or
    /// it already carries revision markup a move would have to re-wrap).
    /// </summary>
    /// <remarks>
    /// One query answers a whole drag: the UI asks on drag start / menu open and gates its drop
    /// indicators and menu items on the result, instead of drawing an indicator over a target the
    /// engine will refuse. <see cref="MoveBlock"/> stays authoritative — this shares its guards
    /// rather than restating them, so a target listed here is one <c>MoveBlock</c> accepts.
    /// Each entry reports the two positions SEPARATELY: a cross-block range or a section break
    /// between the blocks can make one side legal and the other not, so a caller that knows only
    /// "this target is reachable" can still pick the refused side.
    /// </remarks>
    public IReadOnlyList<MoveTarget> ValidMoveTargets(string sourceAnchorId)
    {
        ThrowIfDisposed();
        var empty = Array.Empty<MoveTarget>();

        var sourceTarget = FindAnchor(sourceAnchorId);
        if (sourceTarget is null || sourceTarget.Anchor.Kind is not ("p" or "h" or "li" or "tbl"))
            return empty;
        var source = sourceTarget.Resolve(_doc!);
        if (source?.Parent is not { } parent)
            return empty;
        if (!IsEditorBodyBlock(source, parent))
            return empty;

        // One context for the whole sweep: each candidate then costs index arithmetic over the
        // container's cross-block ranges instead of its own pass over the story.
        var context = BuildBlockMoveContext(parent);
        var run = MovingRun(context, context.Index[source], context.Index[source]);
        if (MoveSourceRejection(context, run) is not null) return empty;

        var targets = new List<MoveTarget>();
        foreach (var candidate in context.Children)
        {
            if (ReferenceEquals(candidate, source) || !IsEditorBodyBlock(candidate, parent))
                continue;
            var unid = (string?)candidate.Attribute(PtOpenXml.Unid);
            if (unid is null) continue;
            var kind = candidate.Name == W.tbl ? "tbl" : WmlToMarkdownConverter.KindFor(candidate);
            if (kind is null) continue;

            int target = context.Index[candidate];
            bool before = LandRun(context, run, target, Position.Before).Error is null;
            bool after = LandRun(context, run, target, Position.After).Error is null;
            if (before || after)
                targets.Add(new MoveTarget($"{kind}:{sourceTarget.Anchor.Scope}:{unid}", before, after));
        }
        return targets;
    }

    /// <summary>The elements whose children are blocks (paragraphs, tables, block content
    /// controls) wherever they appear. A content control's content or a custom XML element holds
    /// blocks only when the control itself sits among blocks.</summary>
    private static readonly HashSet<XName> BlockContainerNames = new()
    {
        W.body, W.tc, W.txbxContent, W.hdr, W.ftr, W.footnote, W.endnote, W.comment, W.docPartBody,
    };

    private static bool HoldsBlocks(XElement? container) => container is not null
        && (BlockContainerNames.Contains(container.Name)
            || (container.Name == W.sdtContent && HoldsBlocks(container.Parent?.Parent))
            || (container.Name == W.customXml && HoldsBlocks(container.Parent)));

    /// <summary>Refuse an anchor that a new paragraph or table cannot be placed beside. Only a
    /// paragraph, a table or a block content control qualifies: a sibling of a cell, row, grid
    /// column, section break, note, comment or a run-, row- or cell-level content control would
    /// be schema-invalid (a <c>w:p</c> directly inside a <c>w:tr</c>, say).</summary>
    private static EditResult? RefuseNonBlockSibling(XElement element, Anchor anchor, string operation) =>
        (element.Name == W.p || element.Name == W.tbl || element.Name == W.sdt) && HoldsBlocks(element.Parent)
            ? null
            : EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"{operation} places a block beside a paragraph, table or block content control; got kind={anchor.Kind}. "
                + "To add content inside a table cell, anchor on one of its paragraphs or use ReplaceCellContent",
                anchor.Id);

    private static bool IsEditorBodyBlock(XElement block, XElement parent)
    {
        if (block.Name != W.p && block.Name != W.tbl) return false;
        if (parent.Name == W.body) return true;
        if (parent.Name != W.sdtContent) return false;
        // The renderer flattens a body-level content control. A content control inside
        // a table cell/text box is not one of the editor's top-level body units.
        return parent.Ancestors(W.body).Any() &&
               !parent.Ancestors(W.tc).Any() &&
               !parent.Ancestors(W.txbxContent).Any();
    }

    /// <summary>
    /// The move-independent facts about one block container: its children in document order, which
    /// of them hold a section break, and where each range and field that reaches from one child
    /// to another starts and ends. <see cref="MoveBlock"/>, <see cref="ValidMoveTargets"/> and
    /// <see cref="MoveBlockRange"/> all decide on it.
    /// </summary>
    /// <remarks>
    /// Everything here is a property of the CONTAINER, not of a particular move, so a whole
    /// <see cref="ValidMoveTargets"/> sweep builds it once and then answers each of its 2N
    /// candidate questions with index arithmetic. Recomputing it per question — five
    /// <c>Descendants</c> passes over the story and a materialized member set per range — made
    /// the sweep quadratic in the block count and linear again in the marker count on top:
    /// on a 234-block charter carrying 392 bookmarks it cost seconds, which the drag UI paid
    /// on every drag start and menu open.
    /// </remarks>
    private sealed class BlockMoveContext
    {
        /// <summary>The container's children: its blocks, and the content controls, range markers
        /// and section properties among them.</summary>
        public required IReadOnlyList<XElement> Children { get; init; }

        /// <summary>Position of each child in <see cref="Children"/>, by reference.</summary>
        public required Dictionary<XElement, int> Index { get; init; }

        /// <summary>Number of section breaks in <c>Children[0..i)</c>, so "is there one between
        /// these two children" is a subtraction.</summary>
        public required int[] SectionBreaksBefore { get; init; }

        /// <summary>Every range and field whose two ends sit in DIFFERENT children, as the index
        /// pair the move has to leave intact; an end outside the container is past its first or
        /// last child. A null label is a pair to keep when possible (see
        /// <see cref="BuildBlockMoveContext"/>).</summary>
        public required IReadOnlyList<(int Start, int End, string? Label)> CrossBlockRanges { get; init; }
        /// <summary>Range markers whose end comes before their start (no range, as ListBookmarks says):
        /// a move may not put the start before the end, which would make one.</summary>
        public required IReadOnlyList<(int End, int Start, string Label)> InvertedPairs { get; init; }

        public IEnumerable<XElement> Run((int First, int Last) run) => Children.Take(run.Last + 1).Skip(run.First);

        /// <summary>The child at <paramref name="i"/> once the run is lifted out.</summary>
        public XElement Other((int First, int Last) run, int i) => Children[i < run.First ? i : i + run.Last - run.First + 1];
    }

    private static BlockMoveContext BuildBlockMoveContext(XElement container)
    {
        var children = container.Elements().ToList();
        // XElement does not override Equals, so the default comparer IS reference identity.
        var index = new Dictionary<XElement, int>(children.Count);
        for (int i = 0; i < children.Count; i++) index[children[i]] = i;
        var sectionBreaksBefore = new int[children.Count + 1];
        var ranges = new List<(int, int, string?)>();
        void Add(int start, int end, string label)
        {
            if (start == end) return;
            ranges.Add((start, end, label));
            if (start < 0 || end == children.Count || !children.Take(end + 1).Skip(start).All(IsBlockLevelMarker)) return;
            // An empty range between blocks marks the start of the block after it: they stay together when they can.
            int next = end;
            while (next < children.Count && IsBlockLevelMarker(children[next])) next++;
            if (next < children.Count) ranges.Add((end, next, null));
        }

        // One traversal finds every marker, with the child it is in: -1 before the container, and past its
        // last child after it. The first start wins for a duplicated id. An end before its start is no range,
        // and a text box is a story of its own, whose fields and ranges pair with nothing outside it.
        var body = container.AncestorsAndSelf(W.body).First();
        var story = container.AncestorsAndSelf(W.txbxContent).FirstOrDefault();
        var inverted = new List<(int, int, string)>();
        var endsFirst = new Dictionary<(XName End, string Id), int>();
        var past = container.AncestorsAndSelf().TakeWhile(e => e != body)
            .Select(e => e.ElementsAfterSelf().FirstOrDefault()).FirstOrDefault(e => e is not null);
        var starts = new Dictionary<(XName End, string Id), int>();
        // Each open field: the child its begin is in, and the child of its separate once seen.
        var fields = new Stack<(int Begin, int? Separate)>();
        int at = -1;
        foreach (var e in body.Descendants())
        {
            if (ReferenceEquals(e.Parent, container)) at = index[e];
            else if (ReferenceEquals(e, past)) at = children.Count;
            if ((e.Name == W.fldChar || CrossBlockRangeNames.ContainsKey(e.Name))
                && e.Ancestors(W.txbxContent).FirstOrDefault() != story)
                continue;
            if (e.Name == W.fldChar)
            {
                var type = (string?)e.Attribute(W.fldCharType);
                if (type == "begin") fields.Push((at, null));
                else if (type == "separate" && fields.Count > 0 && fields.Peek().Separate is null)
                    fields.Push((fields.Pop().Begin, at));
                else if (type == "end" && fields.Count > 0)
                {
                    // The separate is where the result begins: instructions and result are each a range.
                    var (begin, separate) = fields.Pop();
                    if (separate is { } middle && middle != begin && middle != at)
                    {
                        Add(begin, middle, "field");
                        Add(middle, at, "field");
                    }
                    else Add(begin, at, "field");
                }
            }
            else if (e.Name == W.sectPr)
            {
                if (at >= 0 && at < children.Count && e.Parent!.Name == W.pPr && IsStoryParagraph(e.Parent.Parent!))
                    sectionBreaksBefore[at + 1]++;
            }
            else if (CrossBlockRangeNames.TryGetValue(e.Name, out var range))
            {
                var key = (range.End, (string?)e.Attribute(W.id) ?? "");
                if (e.Name != range.End)
                {
                    if (starts.TryAdd(key, at) && endsFirst.TryGetValue(key, out int end) && end >= 0 && end < at && at < children.Count)
                        inverted.Add((end, at, range.Label));
                }
                else if (starts.TryGetValue(key, out int start)) Add(start, at, range.Label);
                else endsFirst.TryAdd(key, at);
            }
        }
        for (int i = 0; i < children.Count; i++) sectionBreaksBefore[i + 1] += sectionBreaksBefore[i];

        return new BlockMoveContext
        {
            Children = children,
            Index = index,
            SectionBreaksBefore = sectionBreaksBefore,
            CrossBlockRanges = ranges,
            InvertedPairs = inverted,
        };
    }

    /// <summary>What moves with the children from <paramref name="first"/> to <paramref name="last"/>:
    /// the markers right beside them (Word ends a bookmark over whole paragraphs there) whose other
    /// ends move too, and an empty range before them. A native move range never moves. A tracked move
    /// takes no markers: both its copies are live, and a marker could go with only one of them.</summary>
    private (int First, int Last) MovingRun(BlockMoveContext context, int first, int last)
    {
        if (_trackedChanges == TrackedChangeMode.RenderInline) return (first, last);
        int lo = first, hi = last;
        while (lo > 0 && IsBlockLevelMarker(context.Children[lo - 1])) lo--;
        while (hi + 1 < context.Children.Count && IsBlockLevelMarker(context.Children[hi + 1])) hi++;

        // The other ends of the ranges that end at a child, the kept-when-possible ones left out.
        IEnumerable<(int Other, string Label)> Ends(int child)
        {
            foreach (var (start, end, label) in context.CrossBlockRanges)
                if (label is not null && (start == child || end == child)) yield return (start == child ? end : start, label);
        }
        bool Goes(int marker) => Ends(marker).All(e =>
            e.Other >= lo && e.Other <= hi && !e.Label.StartsWith("move-", StringComparison.Ordinal));
        // A marker that stays keeps the ones beyond it, which may keep another one's other end.
        for (bool stays = true; stays;)
        {
            stays = false;
            for (int i = first - 1; i >= lo; i--) if (!Goes(i)) { lo = i + 1; stays = true; }
            for (int i = last + 1; i <= hi; i++) if (!Goes(i)) { hi = i - 1; stays = true; }
        }
        while (lo < first && !Ends(lo).Any()) lo++;
        // After the run, only as far as those need: an empty range there belongs to the next block.
        int needed = last;
        for (int i = lo; i <= needed; i++)
            foreach (var (other, _) in Ends(i))
                if (other > needed && other <= hi) needed = other;
        return (lo, needed);
    }

    /// <summary>Where a run of children lands beside <paramref name="target"/>: its slot among the
    /// container's other children, whether that is where it already is, or why it may not land
    /// there: it would cross a section break, or change the membership or order of a cross-block
    /// comment, bookmark, permission, native-move or field range.</summary>
    /// <remarks>
    /// The run can land on either side of the markers between the target and the next block, so
    /// the first such slot that keeps every range is taken. The reordered document order is a function
    /// of the run's bounds and the slot and needs no second list. A range survives iff its two
    /// endpoints still bound a window of the same width: the elements strictly between them can
    /// only change by the run entering or leaving, so equal width and equal membership are the
    /// same condition — including when an endpoint is IN the run, where any relocation but the
    /// identity one changes the width.
    /// </remarks>
    private static (int Slot, bool InPlace, string? Error) LandRun(
        BlockMoveContext context, (int First, int Last) run, int target, Position pos)
    {
        int length = run.Last - run.First + 1, others = context.Children.Count - length;

        // Checked before the ranges, matching the order the messages were produced in
        // when the range scan ran per candidate.
        int from = Math.Min(run.First, target), to = Math.Max(run.Last, target);
        var error = context.SectionBreaksBefore[to + 1] - context.SectionBreaksBefore[from] > 0
            ? "cannot move a block across a section-break paragraph" : null;

        int step = pos == Position.Before ? -1 : 1, best = -1;
        string? firstBroken = null;
        bool bestKeepsAll = false;
        for (int slot = (target > run.Last ? target - length : target) + (pos == Position.Before ? 0 : 1); ; slot += step)
        {
            if (slot == run.First) return (slot, true, error);
            var (required, all) = Keeps(slot);
            if (required is not null) firstBroken ??= required;
            else if (best < 0 || (all && !bestKeepsAll)) (best, bestKeepsAll) = (slot, all);
            int beside = pos == Position.Before ? slot - 1 : slot;
            if (beside < 0 || beside >= others || !IsBlockLevelMarker(context.Other(run, beside))) break;
        }
        return (best, false, error ?? (best < 0 ? $"move would change or invert a cross-block {firstBroken} range" : null));

        (string? Broken, bool All) Keeps(int slot)
        {
            int Reordered(int i)
            {
                if (i >= run.First && i <= run.Last) return slot + (i - run.First);
                int afterRemoval = i > run.Last ? i - length : i;
                return afterRemoval >= slot ? afterRemoval + length : afterRemoval;
            }
            bool all = true;
            foreach (var (start, end, label) in context.CrossBlockRanges)
            {
                if (Reordered(end) - Reordered(start) == end - start) continue;
                if (label is not null) return (label, false);
                all = false;
            }
            foreach (var (end, start, label) in context.InvertedPairs)
                if (Reordered(start) <= Reordered(end)) return (label, false);
            return (null, all);
        }
    }

    /// <summary>Move a run of children to a slot <see cref="LandRun"/> chose, and return the
    /// paragraphs it took to keep tables apart.</summary>
    private List<XElement> RelocateRun(BlockMoveContext context, (int First, int Last) run, int slot)
    {
        XElement first = context.Children[run.First], last = context.Children[run.Last];
        // The run's own siblings between its ends travel with it.
        var nodes = new List<XNode> { first };
        while (!ReferenceEquals(nodes[^1], last)) nodes.Add(nodes[^1].NextNode!);
        var leftBehind = new XElement(W.p);
        first.AddBeforeSelf(leftBehind);
        foreach (var node in nodes) node.Remove();
        InsertAtSlot(context, run, slot, nodes);
        return KeepTablesApart(leftBehind, first, last);
    }

    private static void InsertAtSlot(BlockMoveContext context, (int First, int Last) run, int slot, object content)
    {
        if (slot < context.Children.Count - (run.Last - run.First + 1)) context.Other(run, slot).AddBeforeSelf(content);
        else context.Other(run, slot - 1).AddAfterSelf(content);
    }

    /// <summary>Keep the empty paragraph <paramref name="leftBehind"/>, put where the moved blocks came from, and
    /// add one before <paramref name="first"/> or after <paramref name="last"/> of them, wherever a table
    /// would otherwise touch another one or end the body (as <see cref="InsertTable"/> does): tables
    /// that touch are read as one, and Word always has a paragraph after the last one. Returns the
    /// ones kept.</summary>
    private List<XElement> KeepTablesApart(XElement? leftBehind, XElement first, XElement last)
    {
        var kept = new List<XElement>();
        foreach (var place in new Action<XElement>?[] { null, first.AddBeforeSelf, last.AddAfterSelf })
        {
            if (place is null && leftBehind is null) continue;
            var separator = place is null ? leftBehind! : new XElement(W.p);
            place?.Invoke(separator);
            PlaceOutsideRanges(separator);
            // The tables stay apart in the markup as it stands and once the edit is accepted; an edit that is
            // not tracked has to hold once the other pending changes are rejected too.
            if (TablesMeet(separator, accepted: true) || TablesMeet(separator, accepted: null)
                || (_trackedChanges != TrackedChangeMode.RenderInline && TablesMeet(separator, accepted: false)))
            {
                UnidHelper.AssignToSelfAndDescendants(separator);
                kept.Add(separator);
            }
            else separator.Remove();
        }
        return kept;
    }

    /// <summary>Move <paramref name="separator"/> within the markers around it (between the blocks on either side) to
    /// a place inside the fewest of the ranges that start or end there, staying where it is on a tie: past the ends
    /// of those that cover the table, short of those that start after it. A range with no marker there contains every
    /// such place or none. An empty range keeps covering nothing, so where no place outside the others is left before
    /// it, it can come to mark the added paragraph rather than the table after it. A content control's revision
    /// markers stay against the control they enclose: the paragraph never goes between them.</summary>
    private static void PlaceOutsideRanges(XElement separator)
    {
        static XElement? Sibling(XElement e, bool forward)
        {
            for (var node = forward ? e.NextNode : e.PreviousNode; node is not null; node = forward ? node.NextNode : node.PreviousNode)
                if (node is XElement element) return element;
            return null;
        }
        var markers = new List<XElement>();
        for (var e = Sibling(separator, forward: false); e is not null && IsBlockLevelMarker(e); e = Sibling(e, forward: false))
            markers.Add(e);
        markers.Reverse();
        int here = markers.Count;
        for (var e = Sibling(separator, forward: true); e is not null && IsBlockLevelMarker(e); e = Sibling(e, forward: true))
            markers.Add(e);

        // How many of those ranges each place among the markers is inside, counted in one pass: an end is
        // inside from its start (or from before the markers) up to it, a start whose end is further on after it.
        int count = markers.Count, lo = 0, hi = count;
        var inside = new int[count + 2];
        var open = new Dictionary<(XName End, string Id), int>();
        (XName End, string Id)? Key(XElement m) =>
            CrossBlockRangeNames.TryGetValue(m.Name, out var r) ? (r.End, (string?)m.Attribute(W.id) ?? "") : null;
        // A pair with both ends here encloses no block, so it holds no control to stay against.
        var keys = markers.Select(Key).ToList();
        var tally = keys.OfType<(XName End, string Id)>().GroupBy(k => k).ToDictionary(g => g.Key, g => g.Count());
        for (int i = 0; i < count; i++)
        {
            if (keys[i] is not { } key) continue;
            bool isEnd = markers[i].Name == key.End;
            if (markers[i].Name.LocalName.StartsWith("customXml", StringComparison.Ordinal) && tally[key] == 1)
            {
                if (isEnd) lo = Math.Max(lo, i + 1);
                else hi = Math.Min(hi, i);
            }
            if (!isEnd) { open.TryAdd(key, i); continue; }
            inside[open.Remove(key, out int from) ? from + 1 : 0]++;
            inside[i + 1]--;
        }
        foreach (var from in open.Values) { inside[from + 1]++; inside[count + 1]--; }
        for (int slot = 1; slot <= count; slot++) inside[slot] += inside[slot - 1];
        if (lo > hi) return;

        int best = here >= lo && here <= hi ? here : lo;
        for (int slot = lo; slot <= hi; slot++)
            if (inside[slot] < inside[best]) best = slot;
        if (best == here) return;
        separator.Remove();
        if (best == 0) markers[0].AddBeforeSelf(separator);
        else markers[best - 1].AddAfterSelf(separator);
    }

    /// <summary>Whether two tables would touch, or a table end the body, without <paramref name="separator"/>:
    /// a table before it and a table, the section properties or nothing after it. Resolving a revision
    /// asks this of the document as it stands (<paramref name="accepted"/> null) before it removes an
    /// empty paragraph (<see cref="Internal.RevisionOps.ClosesTable"/>).</summary>
    internal static bool TablesMeet(XElement separator, bool? accepted) =>
        NeighbourBlock(separator, forward: false, accepted)?.Name == W.tbl
        && NeighbourBlock(separator, forward: true, accepted)?.Name is var next
        && (next is null || next == W.sectPr || next == W.tbl);

    /// <summary>The block before or after <paramref name="block"/> once the pending changes are accepted
    /// or rejected (or as they stand, when <paramref name="accepted"/> is null): past markers, the edges
    /// of block content controls (neither keeps two tables apart) and blocks that go. Null at the
    /// body's start or end.</summary>
    private static XElement? NeighbourBlock(XElement block, bool forward, bool? accepted)
    {
        for (XElement? e = block; e is not null; e = Internal.RevisionOps.BlockWrapperOf(e))
        {
            var beside = WmlToMarkdownConverter.ExpandBlockWrappers(forward ? e.ElementsAfterSelf() : e.ElementsBeforeSelf());
            if ((forward ? beside : beside.Reverse())
                .FirstOrDefault(b => !IsBlockLevelMarker(b) && b.Name != W.customXmlPr && !(accepted is { } view && Goes(b, view))) is { } found)
                return found;
        }
        return null;
    }

    /// <summary>Whether a paragraph with its mark, or a table, goes when every pending change is accepted
    /// (a deletion or move source) or rejected (an insertion or move destination).</summary>
    private static bool Goes(XElement block, bool accepted)
    {
        XName revised = accepted ? W.del : W.ins, moved = accepted ? W.moveFrom : W.moveTo;
        bool IsMark(XElement e) => e.Name == revised || e.Name == moved;
        return block.Name == W.tbl
            ? WordprocessingMLUtil.TableRows(block).All(row => row.Element(W.trPr)?.Element(revised) is not null)
            : block.Element(W.pPr)?.Element(W.rPr)?.Elements().Any(IsMark) == true
                && block.Descendants().Where(e => e.Name == W.r || e.Name == M.r).All(r => r.Ancestors().Any(IsMark));
    }

    /// <summary>A range marker (or a proofing mark) that is itself a child of the body or of a block
    /// content control, so not a block.</summary>
    internal static bool IsBlockLevelMarker(XElement child) =>
        child.Name == W.proofErr || CrossBlockRangeNames.ContainsKey(child.Name);

    /// <summary>
    /// Wrap the live runs of <paramref name="paragraph"/> in a
    /// <paramref name="wrapperName"/> revision envelope and optionally mark the paragraph mark to match.
    /// </summary>
    /// <remarks>
    /// Shared by whole-block deletion, text replacement, paragraph/table moves and paragraph/row insertion.
    /// A deletion also converts
    /// <c>w:t</c>→<c>w:delText</c> (and <c>w:instrText</c>→<c>w:delInstrText</c>), which is what
    /// Word writes, what the comparison renderer produces, and what
    /// <see cref="RevisionProcessor"/>'s reject path swaps back. Move sources retain ordinary text.
    /// The paragraph mark belongs
    /// to the same operation unless the caller retains it (a replacement or a final retained paragraph).
    /// </remarks>
    private XElement? MarkParagraphContentAndMark(
        XElement paragraph, XName wrapperName, string author, string date,
        bool preserveParagraphMark = false, bool retainInlineStructure = false)
    {
        // Hyperlinks/fields own their revised runs. A deletion also owns math objects (a
        // w:del > m:oMath envelope, as the comparison renderer writes) and, when the block
        // itself goes, inline SDTs: an inline control can itself be a w:del child, and deleting
        // its existence avoids leaving an empty control in the following paragraph after its
        // pilcrow is accepted. A content replacement keeps the control shell and deletes inside
        // it instead, so a run another author inserted there is deleted BELOW that insertion —
        // the shape Word writes, and the only one whose deletion rejects on its own.
        var stamp = new RevisionStamp(author, date);
        var content = paragraph.Descendants().Where(e => e.Name == W.r
            || (wrapperName == W.del && (e.Name == M.oMath || e.Name == M.oMathPara
                || (!retainInlineStructure && e.Name == W.sdt && IsInlineControl(e))))).ToList();
        foreach (var element in content)
        {
            if (wrapperName == W.del)
            {
                DeleteInlineElementInPlace(element, stamp, keepMarkers: retainInlineStructure);
                continue;
            }
            // Existing deletions/move sources are already absent from the accepted view;
            // ordinary insertion/move marking retains its existing revision ownership.
            if (element.Ancestors().Any(e => e.Name == W.del || e.Name == W.moveFrom
                    || e.Name == W.ins || e.Name == W.moveTo))
                continue;
            var envelope = CreateRevisionEnvelope(wrapperName, stamp);
            element.ReplaceWith(envelope);
            envelope.Add(element);
        }

        return preserveParagraphMark
            ? paragraph.Element(W.pPr)
            : MarkParagraphMark(paragraph, wrapperName, author, date);
    }

    /// <summary>Mark <paramref name="paragraph"/>'s own paragraph mark with a
    /// <paramref name="wrapperName"/> revision (<c>w:pPr/w:rPr/w:ins</c> and its siblings),
    /// unless it already carries one, and return its <c>w:pPr</c>.</summary>
    private XElement MarkParagraphMark(XElement paragraph, XName wrapperName, string author, string date)
    {
        var pPr = paragraph.Element(W.pPr);
        if (pPr is null)
        {
            pPr = new XElement(W.pPr);
            paragraph.AddFirst(pPr);
        }
        var rPr = GetOrCreatePPrChild(pPr, W.rPr);
        if (rPr.Element(wrapperName) is null)
            WordprocessingMLUtil.InsertRPrChildInOrder(rPr, CreateRevisionEnvelope(wrapperName, author, date));
        return pPr;
    }

    /// <summary>
    /// Delete one run-level element in place the way Word does, keeping the hyperlink, field
    /// and control shells around it. An ordinary run becomes <c>w:del</c>; a run already deleted
    /// or moved away stays as it is; a run inside another author's insertion is deleted inside
    /// that insertion (<c>w:ins &gt; w:del</c>, so the deletion rejects on its own); a run inside
    /// the session author's own insertion is simply un-inserted — retyping your own pending text
    /// leaves no struck-through first attempt — except a zero-width marker run (a note or comment
    /// reference), which is only ever marked so a reject can still restore it. With
    /// <paramref name="keepMarkers"/> a marker-only run is left live: a content replacement
    /// supersedes text, not the references that must survive it on accept and reject. The one
    /// owner of this policy for whole-paragraph, block and control-fill deletions.
    /// </summary>
    private void DeleteInlineElementInPlace(XElement element, RevisionStamp stamp, bool keepMarkers)
    {
        bool marker = IsMarkerOnlyRun(element);
        if (keepMarkers && marker) return;
        if (element.Ancestors().Any(e => e.Name == W.del || e.Name == W.moveFrom)) return;
        if (!marker && OwnPendingInsertion(element, stamp.Author) is { } insertion)
        {
            RemoveFromInsertion(element, insertion);
            return;
        }
        var envelope = CreateRevisionEnvelope(W.del, stamp);
        element.ReplaceWith(envelope);
        envelope.Add(element);
        RevisionOps.SpellTextAsDeleted(element);
    }

    /// <summary>Remove <paramref name="element"/> from the session author's own pending
    /// <paramref name="insertion"/>, together with every container it leaves empty on the way up
    /// to the paragraph: the envelope itself, an enclosing envelope, and a hyperlink or simple
    /// field with nothing else in it (an empty field is a shape the tracked deleter refuses). An
    /// emptied control keeps its shell unless the control itself sat inside the insertion — was
    /// the author's own — and a move envelope is never removed, its partner still needs it.</summary>
    private static void RemoveFromInsertion(XElement element, XElement insertion)
    {
        var parent = element.Parent;
        element.Remove();
        while (parent is not null && parent.Name != W.p && !parent.HasElements)
        {
            var emptied = parent;
            parent = emptied.Parent;
            if (emptied.Name == W.moveTo || emptied.Name == W.moveFrom) break;
            if (emptied.Name == W.sdtContent)
            {
                if (parent is null || parent.Name != W.sdt
                    || !parent.Ancestors().Any(a => ReferenceEquals(a, insertion)))
                    break;
                var control = parent;
                parent = control.Parent;
                control.Remove();
                continue;
            }
            emptied.Remove();
        }
        // An insertion left holding only range markers (a bookmark it wrapped) would list as an
        // insertion with no text: unwrap it, keeping the markers where they were.
        if (parent is { Parent: not null } && parent.Name == W.ins && parent.Elements().All(IsContentless))
        {
            var kept = parent.Elements().ToList();
            foreach (var child in kept) child.Remove();
            parent.ReplaceWith(kept);
        }
    }

    private bool IsOwnInsertion(XElement? insertion, string author) =>
        insertion is not null
        && IsOwnRevisionAuthor((string?)insertion.Attribute(W.author), author);

    /// <summary>The <c>w:ins</c> that most closely encloses <paramref name="element"/> when it is
    /// <paramref name="author"/>'s own; null when the nearest insertion-like ancestor is another
    /// author's or a move destination (which is never un-inserted), or there is none.</summary>
    private XElement? OwnPendingInsertion(XElement element, string author) =>
        element.Ancestors().FirstOrDefault(e => e.Name == W.ins || e.Name == W.moveTo) is { } insertion
        && insertion.Name == W.ins && IsOwnInsertion(insertion, author)
            ? insertion
            : null;

    /// <summary>Whether a tracked deletion of <paramref name="root"/> removes content outright
    /// rather than marking it: a run (other than a marker-only one) inside the session author's
    /// own pending insertion, which <see cref="DeleteInlineElementInPlace"/> un-inserts.</summary>
    private bool HasOwnInsertedContent(XElement root, string author) =>
        root.Descendants(W.r).Any(run => !IsMarkerOnlyRun(run)
            && !run.Ancestors().Any(a => a.Name == W.del || a.Name == W.moveFrom)
            && OwnPendingInsertion(run, author) is not null);

    /// <summary>Whether <paramref name="paragraph"/> is, in its entirety, the session author's
    /// own pending insertion: its pilcrow carries the author's <c>w:ins</c> and every child is
    /// the author's insertion (directly, or through a hyperlink, field or smart tag holding only
    /// the author's insertions), with no range marker whose partner may lie elsewhere. Such a
    /// paragraph is un-inserted outright by a tracked deletion, as Word does.</summary>
    private bool IsWhollyOwnInsertion(XElement paragraph, string author) =>
        IsOwnInsertion(paragraph.Element(W.pPr)?.Element(W.rPr)?.Element(W.ins), author)
        && paragraph.Elements().Where(child => child.Name != W.pPr).All(child => IsOwnInsertedContent(child, author))
        && !paragraph.Descendants().Any(d => d.Name == W.bookmarkStart || d.Name == W.bookmarkEnd
            || d.Name == W.commentRangeStart || d.Name == W.commentRangeEnd
            || d.Name == W.permStart || d.Name == W.permEnd
            || d.Name == W.moveFromRangeStart || d.Name == W.moveFromRangeEnd
            || d.Name == W.moveToRangeStart || d.Name == W.moveToRangeEnd);

    private bool IsOwnInsertedContent(XElement element, string author) =>
        element.Name == W.proofErr
        || (element.Name == W.ins && IsOwnInsertion(element, author))
        || ((element.Name == W.hyperlink || element.Name == W.fldSimple || element.Name == W.smartTag)
            && element.HasElements && element.Elements().All(child => IsOwnInsertedContent(child, author)));

    private void MarkParagraphAsTrackedMove(
        XElement paragraph, bool from, string moveName, string author, string date)
    {
        EnsureTrackRevisionsEnabled();
        var pPr = MarkParagraphContentAndMark(
            paragraph, from ? W.moveFrom : W.moveTo, author, date)!;

        int rangeId = NextRevisionId();
        var start = new XElement(from ? W.moveFromRangeStart : W.moveToRangeStart,
            new XAttribute(W.id, rangeId),
            new XAttribute(W.name, moveName),
            new XAttribute(W.author, author),
            new XAttribute(W.date, date));
        var end = new XElement(from ? W.moveFromRangeEnd : W.moveToRangeEnd,
            new XAttribute(W.id, rangeId));
        pPr.AddAfterSelf(start);
        paragraph.Add(end);
    }

    /// <summary>
    /// Mark a whole table as inserted or deleted: the row-existence revision on every
    /// <c>w:trPr</c> AND the cell content, matching the design's whole-table lowering ("every
    /// source row <em>and its content</em> is deleted"). Row marks alone leave the moved-away
    /// table's text rendering as ordinary body text inside a row Word believes is deleted.
    /// </summary>
    private void MarkTableRowsAsTrackedRevision(
        XElement table, bool inserted, string author, string date)
    {
        EnsureTrackRevisionsEnabled();
        foreach (var row in table.Descendants(W.tr).ToList())
            MarkRowAsTrackedRevision(row, inserted, author, date);
    }

    public EditResult InsertParagraph(string anchorId, Position pos, string markdownPayload)
    {
        if (MutationRefusal() is { } refusal) return EditResult.Fail(refusal);
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);

        var parsed = Internal.MarkdownPayloadParser.Parse(markdownPayload);
        if (!parsed.Success)
            return EditResult.Fail(parsed.Error!.Code, parsed.Error.Message, anchorId);
        if (ValidatePendingHyperlinks(parsed.Blocks.SelectMany(b => b.RunElements), anchorId) is { } linkError)
            return linkError;
        if (parsed.Blocks.Count == 0)
            return EditResult.Fail(EditErrorCode.MalformedMarkdown, "empty payload", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);
        if (RefuseNonBlockSibling(element, target.Anchor, "InsertParagraph") is { } notBlock) return notBlock;

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var created = new List<Anchor>();
            var newElements = new List<XElement>();
            var lists = new PayloadListState(parsed.Blocks)
            {
                PrecedingNeighbor = pos == Position.After
                    ? element
                    : element.ElementsBeforeSelf().LastOrDefault(),
                FollowingNeighbor = pos == Position.After
                    ? element.ElementsAfterSelf().FirstOrDefault()
                    : element,
            };
            foreach (var block in parsed.Blocks)
            {
                var p = BuildParagraphFromParsedBlock(block);
                AssignPayloadListNumbering(p, block, lists);
                UnidHelper.AssignToSelfAndDescendants(p);
                newElements.Add(p);
                var unid = (string)p.Attribute(PtOpenXml.Unid)!;
                var kind = ClassifyParagraphKind(p);
                created.Add(new Anchor($"{kind}:{target.Anchor.Scope}:{unid}", kind, target.Anchor.Scope, unid));
            }

            if (pos == Position.Before)
            {
                foreach (var n in newElements) element.AddBeforeSelf(n);
            }
            else
            {
                XElement after = element;
                foreach (var n in newElements) { after.AddAfterSelf(n); after = n; }
            }

            if (_trackedChanges == TrackedChangeMode.RenderInline)
            {
                var author = _revisionAuthor ?? "docxodus";
                var date = NextTrackedFormatRevisionDate();
                foreach (var paragraph in newElements)
                    MarkParagraphContentAndMark(paragraph, W.ins, author, date);
            }

            foreach (var n in newElements) PromoteHyperlinkRelationships(n);

            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Created = created,
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, anchorId);
        }
    }

    /// <summary>The paragraph mark's own insertion, deletion or move markers
    /// (<c>w:pPr/w:rPr/w:ins</c> and its siblings).</summary>
    private static IEnumerable<XElement> ParagraphMarkRevisions(XElement pPr) =>
        pPr.Elements(W.rPr).Elements().Where(e => Internal.RevisionOps.RevWrapperNames.Contains(e.Name));

    /// <summary>Drop what belongs to the paragraph mark alone from <paramref name="pPr"/>: the
    /// <see cref="ParagraphMarkRevisions"/>, then the mark's <c>w:rPr</c> when nothing else is
    /// left in it, and the section break (<c>w:sectPr</c>) the mark ends.</summary>
    private static void RemoveMarkOwnedProperties(XElement pPr)
    {
        ParagraphMarkRevisions(pPr).Remove();
        if (pPr.Element(W.rPr) is { HasElements: false } markProperties) markProperties.Remove();
        pPr.Elements(W.sectPr).Remove();
    }

    /// <summary>
    /// Rejecting a tracked split joins the first half into the second, which keeps the original
    /// paragraph mark and so keeps its own properties. Record whatever the split changed there
    /// (the page break left on the first half, the next style a heading's Enter starts, a mark
    /// formatting the fresh paragraph drops) as this edit's property changes, so rejecting
    /// everything gives back the original paragraph's properties. A pending change copied from
    /// the original keeps its baseline, which is what a reject restores anyway, and is credited
    /// to this edit like any other fold.
    /// </summary>
    private void TrackSplitOffProperties(XElement originalPPr, XElement secondPPr, RevisionStamp stamp)
    {
        var markRevisions = Internal.RevisionOps.RevWrapperNames;
        var mark = secondPPr.Element(W.rPr) ?? new XElement(W.rPr);
        TrackPropertyEdit(mark, W.rPrChange,
            PropertySnapshot(originalPPr.Element(W.rPr), W.rPr, W.rPrChange, markRevisions), stamp, markRevisions);
        // The mark's rPr is attached only when it holds something: a resolved change can empty it.
        if (mark.Parent is null && mark.HasElements) GetOrCreatePPrChild(secondPPr, W.rPr).Add(mark.Elements());
        else if (mark.Parent is not null && !mark.HasElements) mark.Remove();
        TrackPropertyEdit(secondPPr, W.pPrChange,
            PropertySnapshot(originalPPr, W.pPr, W.pPrChange, W.rPr, W.sectPr), stamp, W.rPr, W.sectPr);
    }

    public EditResult SplitParagraph(string anchorId, int characterOffset)
    {
        if (MutationRefusal() is { } refusal) return EditResult.Fail(refusal);
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "SplitParagraph requires a paragraph anchor", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);

        var totalText = ParagraphText(element);
        if (characterOffset < 0 || characterOffset > totalText.Length)
            return EditResult.Fail(EditErrorCode.OffsetOutOfRange,
                $"offset {characterOffset} out of [0, {totalText.Length}]", anchorId);
        // MoveInlineChildrenAfter relocates whole top-level children, so an offset strictly
        // inside an atomic container would silently split at that container's far edge and
        // still report success. Refuse instead — same boundary contract as note citations.
        if (HasUnsupportedInlineInsertionBoundary(element, characterOffset))
            return EditResult.Fail(EditErrorCode.UnsupportedInlineBoundary,
                "SplitParagraph offset falls inside a revision or unsupported inline container; " +
                "choose a boundary before or after that container", anchorId);
        // A field whose begin and end are both here stays whole: an offset between its begin and
        // end marks would leave one in each paragraph. Its edges stay splittable.
        if (SplitWouldCutAField(element, characterOffset))
            return EditResult.Fail(EditErrorCode.UnsupportedInlineBoundary,
                "SplitParagraph offset falls between a field's begin and end marks; split before or after the field",
                anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var pPr = element.Element(W.pPr);
            var second = new XElement(W.p);
            XElement? newPPr = null;
            if (pPr is not null)
            {
                newPPr = new XElement(pPr);
                second.Add(newPPr);
                // Both halves keep any pending property change, each under its own id. The new
                // mark ends the first half, as Word's Enter does, and the original mark goes on
                // ending the second, so a revision of the mark stays there alone, and so do the
                // original ids: the new mark of an untracked split is ordinary content, and a
                // copy of the revision on it would let a reject also take back the split
                // (issue #1297). The section break the mark ends stays with it too: a copy
                // would add a section (issue #1298).
                RemoveMarkOwnedProperties(pPr);
                Internal.RevisionOps.RenumberCopiedRevisions(pPr, NextRevisionId);
            }

            // Split any run that straddles the offset (descends into hyperlinks/sdts),
            // then split any container (hyperlink) that still straddles, then move all
            // inline children + markers at-or-past the offset to `second`.
            SplitParagraphContent(element, characterOffset, second, NextRevisionId);

            if (newPPr is not null)
            {
                // pageBreakBefore is a once-only property: the original paragraph keeps it; the new
                // paragraph must not inherit a second page break (matches Word clearing it on Enter).
                newPPr.Elements(W.pageBreakBefore).Remove();

                // An empty bordered paragraph is a horizontal rule; splitting it (Enter) must not
                // propagate the rule's border onto the fresh paragraph below — otherwise every Enter
                // stacks another rule and borders the body text (S-1 smoke-test finding 1a). A bordered
                // paragraph that HAS text keeps its border on both halves (boxed-block behavior).
                if (totalText.Length == 0)
                    newPPr.Elements(W.pBdr).Remove();

                // An empty Enter-at-end split starts a fresh paragraph. For a non-list paragraph whose
                // style declares a distinct next-paragraph style (e.g. Title/Heading -> Normal), rebase
                // the new paragraph onto that next style instead of cloning the heading: a clean pStyle,
                // dropping the heading-only direct props and the inherited paragraph-mark rPr that would
                // otherwise bake the heading's bold into freshly-typed text. List items are exempt so the
                // editor's Enter-continuation keeps the list going.
                bool emptySplit = characterOffset >= totalText.Length;
                bool isListItem = newPPr.Element(W.numPr) is not null;
                if (emptySplit && !isListItem)
                {
                    var curStyle = (string?)newPPr.Element(W.pStyle)?.Attribute(W.val);
                    var nextStyle = ResolveNextParagraphStyle(curStyle);
                    if (nextStyle is not null && !string.Equals(nextStyle, curStyle, StringComparison.Ordinal))
                    {
                        // What the mark owns comes along. Under tracking, so do the pending
                        // changes a reject must still be able to undo: TrackSplitOffProperties
                        // then keeps their baselines instead of recording the heading as one.
                        var tracked = _trackedChanges == TrackedChangeMode.RenderInline;
                        var markProperties = ParagraphMarkRevisions(newPPr)
                            .Concat(tracked ? newPPr.Elements(W.rPr).Elements(W.rPrChange) : Enumerable.Empty<XElement>())
                            .ToList();
                        var rebuilt = new XElement(W.pPr,
                            new XElement(W.pStyle, new XAttribute(W.val, nextStyle)),
                            markProperties.Count > 0 ? new XElement(W.rPr, markProperties) : null,
                            newPPr.Element(W.sectPr),
                            tracked ? newPPr.Element(W.pPrChange) : null);
                        newPPr.ReplaceWith(rebuilt);
                        newPPr = rebuilt;
                    }
                }

                // Re-mint Unids on the new paragraph's property subtree so cloned property elements
                // (jc, ind, numPr, ...) don't carry the original's Unid onto a second element.
                foreach (var el in newPPr.DescendantsAndSelf())
                    el.Attributes(PtOpenXml.Unid).Remove();
            }

            if (_trackedChanges == TrackedChangeMode.RenderInline)
            {
                // Word's tracked Enter: the new mark that ends the first half is an insertion.
                var stamp = NewRevisionStamp();
                if (pPr is not null && newPPr is not null)
                    TrackSplitOffProperties(pPr, newPPr, stamp);
                MarkParagraphMark(element, W.ins, stamp.Author, stamp.Date);
            }

            UnidHelper.AssignToSelfAndDescendants(second);
            element.AddAfterSelf(second);

            var secondUnid = (string)second.Attribute(PtOpenXml.Unid)!;
            InvalidateProjectionCache();

            // The new paragraph's kind can differ from the original (Heading -> Normal via the
            // next-paragraph style), so derive it from the element itself with the projector's
            // own KindFor — the same derivation a full index rebuild would apply, minus the
            // whole-document walk that made Enter the second-most expensive part of a split.
            // `second` is already attached, so KindFor's style-chain lookup resolves.
            var secondKind = WmlToMarkdownConverter.KindFor(second) ?? target.Anchor.Kind;
            var secondAnchor = new Anchor(
                $"{secondKind}:{target.Anchor.Scope}:{secondUnid}",
                secondKind, target.Anchor.Scope, secondUnid);

            return new EditResult
            {
                Success = true,
                Modified = new[] { target.Anchor },
                Created = new[] { secondAnchor },
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, anchorId);
        }
    }

    /// <summary>
    /// The linked next-paragraph style (<c>w:style/w:next/@w:val</c>) for the given paragraph style
    /// id, read from the styles part; null when the id is empty/unknown or declares no next style.
    /// Read via <c>GetXDocument</c> (the same view <see cref="Internal.StyleFactory"/> writes through)
    /// so styles synthesized earlier in the session are visible.
    /// </summary>
    private string? ResolveNextParagraphStyle(string? styleId)
    {
        if (string.IsNullOrEmpty(styleId)) return null;
        var part = _doc?.MainDocumentPart?.StyleDefinitionsPart;
        var root = part?.GetXDocument().Root;
        if (root is null) return null;
        var style = root.Elements(W.style)
            .FirstOrDefault(st => (string?)st.Attribute(W.styleId) == styleId);
        return (string?)style?.Element(W.next)?.Attribute(W.val);
    }

    /// <summary>
    /// Why a tracked merge of <paramref name="first"/> into the paragraph after it cannot be
    /// recorded, or null when it can (issue #1313):
    /// <list type="bullet">
    /// <item>the deleted mark that records the join needs two paragraphs;</item>
    /// <item>a mark another revision already deleted or moved, on either paragraph, is that
    /// revision's to resolve first: a deleted second mark would join a third paragraph on accept;</item>
    /// <item>a first mark that ends a section owns the section break, which accepting the deletion
    /// would lose where an untracked merge keeps it;</item>
    /// <item>a pending property change on the first paragraph (its <c>w:pPrChange</c> or its
    /// mark's <c>w:rPrChange</c>, by anyone): the merge copies the first paragraph's properties
    /// onto the second, and one copy cannot follow both outcomes of that change, so rejecting it
    /// while accepting the merge would bring back the formatting it rejected;</item>
    /// <item>a pending property change on the second paragraph, unless it is the editor's own
    /// and the merge takes it back to its baseline (Enter then Backspace), which resolves it.
    /// Otherwise the merge would fold into it and the one change would record both edits, so
    /// neither could be reviewed on its own; and resolving another author's change would discard
    /// it, so rejecting the merge would not bring it back.</item>
    /// </list>
    /// </summary>
    private EditResult? RefuseTrackedMerge(XElement first, XElement second, string anchorId)
    {
        var firstPPr = first.Element(W.pPr);
        var secondPPr = second.Element(W.pPr);
        if (firstPPr is not null && ParagraphMarkRevisions(firstPPr).FirstOrDefault(r => r.Name != W.ins) is { } firstMark)
            return UnresolvedBeforeMerge($"the first paragraph's mark has an unresolved {firstMark.Name.LocalName} revision", anchorId);
        if (secondPPr is not null && ParagraphMarkRevisions(secondPPr)
                .FirstOrDefault(r => r.Name == W.del || r.Name == W.moveFrom) is { } secondMark)
            return UnresolvedBeforeMerge($"the second paragraph's mark has an unresolved {secondMark.Name.LocalName} revision", anchorId);
        if (firstPPr?.Element(W.sectPr) is not null)
            return TrackedStructureUnsupported("MergeParagraphs", anchorId);
        if ((firstPPr?.Element(W.pPrChange) ?? firstPPr?.Element(W.rPr)?.Element(W.rPrChange)) is { } firstChange)
            return UnresolvedBeforeMerge($"the first paragraph has an unresolved {firstChange.Name.LocalName}", anchorId);

        var author = _revisionAuthor ?? "docxodus";
        var markRevisions = Internal.RevisionOps.RevWrapperNames;
        bool Resolves(XElement change, XName propertyName, XElement? target, params XName[] excluded) =>
            IsOwnRevisionAuthor((string?)change.Attribute(W.author), author)
            && XNode.DeepEquals(
                PropertySnapshot(change.Element(propertyName), propertyName, change.Name, excluded),
                PropertySnapshot(target, propertyName, change.Name, excluded));
        if (secondPPr?.Element(W.pPrChange) is { } secondChange
            && !Resolves(secondChange, W.pPr, firstPPr, W.rPr, W.sectPr))
            return UnresolvedBeforeMerge("the second paragraph has an unresolved pPrChange", anchorId);
        if (secondPPr?.Element(W.rPr)?.Element(W.rPrChange) is { } secondMarkChange
            && !Resolves(secondMarkChange, W.rPr, firstPPr?.Element(W.rPr), markRevisions))
            return UnresolvedBeforeMerge("the second paragraph's mark has an unresolved rPrChange", anchorId);
        return null;
    }

    private static EditResult UnresolvedBeforeMerge(string what, string anchorId) =>
        EditResult.Fail(EditErrorCode.UnresolvedStructuralRevision,
            $"{what}; resolve it before a tracked MergeParagraphs", anchorId);

    /// <summary>
    /// Accepting a deleted paragraph mark joins its paragraph into the next one, which keeps its
    /// own mark and so its own properties. An untracked merge keeps the first paragraph's other
    /// properties and the second's section boundary, so a
    /// tracked merge gives the second paragraph the first's properties (its mark formatting
    /// included) as this edit's property changes: accepting everything gives what the untracked
    /// merge gives, and rejecting gives the second paragraph its own back. What belongs to the
    /// second mark alone — its own revision and its section break — stays. The mirror of
    /// <see cref="TrackSplitOffProperties"/>.
    /// </summary>
    private void TrackMergedProperties(XElement? firstPPr, XElement second, RevisionStamp stamp)
    {
        var markRevisions = Internal.RevisionOps.RevWrapperNames;
        var pPr = second.Element(W.pPr);
        if (pPr is null)
        {
            pPr = new XElement(W.pPr);
            second.AddFirst(pPr);
        }

        var mark = pPr.Element(W.rPr) ?? new XElement(W.rPr);
        var markBefore = PropertySnapshot(mark, W.rPr, W.rPrChange, markRevisions);
        var markSource = PropertySnapshot(firstPPr?.Element(W.rPr), W.rPr, W.rPrChange, markRevisions);
        Internal.RevisionOps.RenumberCopiedRevisions(markSource, NextRevisionId);
        mark.Elements().Where(e => e.Name != W.rPrChange && !markRevisions.Contains(e.Name)).Remove();
        // CT_ParaRPr: the mark's own revisions, then its run properties, then rPrChange.
        if (mark.Elements().LastOrDefault(e => markRevisions.Contains(e.Name)) is { } lastRevision)
            lastRevision.AddAfterSelf(markSource.Elements());
        else
            mark.AddFirst(markSource.Elements());
        TrackPropertyEdit(mark, W.rPrChange, markBefore, stamp, markRevisions);
        if (mark.Parent is null && mark.HasElements) GetOrCreatePPrChild(pPr, W.rPr).Add(mark.Elements());
        else if (mark.Parent is not null && !mark.HasElements) mark.Remove();

        var before = PropertySnapshot(pPr, W.pPr, W.pPrChange, W.rPr, W.sectPr);
        var source = PropertySnapshot(firstPPr, W.pPr, W.pPrChange, W.rPr, W.sectPr);
        Internal.RevisionOps.RenumberCopiedRevisions(source, NextRevisionId);
        pPr.Elements().Where(e => e.Name != W.rPr && e.Name != W.sectPr && e.Name != W.pPrChange).Remove();
        pPr.AddFirst(source.Elements());
        TrackPropertyEdit(pPr, W.pPrChange, before, stamp, W.rPr, W.sectPr);
        if (!pPr.HasElements) pPr.Remove();
    }

    /// <summary>
    /// Join the paragraph <paramref name="secondAnchorId"/> names onto the end of the one
    /// <paramref name="firstAnchorId"/> names, its immediate previous sibling, with a separating
    /// space when neither side has one. The joined paragraph keeps the first paragraph's
    /// properties except for its section boundary, which comes from the second paragraph.
    /// Under <see cref="TrackedChangeMode.RenderInline"/> the join is recorded as Word
    /// records it: the first paragraph's mark is deleted, both paragraphs stay until review (so
    /// both are in <see cref="EditResult.Modified"/>), and accepting gives the untracked result
    /// (issue #1313).
    /// </summary>
    public EditResult MergeParagraphs(string firstAnchorId, string secondAnchorId)
    {
        if (MutationRefusal() is { } refusal) return EditResult.Fail(refusal);
        var firstTarget = FindAnchor(firstAnchorId);
        if (firstTarget is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "first anchor not found", firstAnchorId);
        var secondTarget = FindAnchor(secondAnchorId);
        if (secondTarget is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "second anchor not found", secondAnchorId);

        var firstEl = firstTarget.Resolve(_doc!);
        var secondEl = secondTarget.Resolve(_doc!);
        if (firstEl is null || secondEl is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null");

        if (firstEl.Name != W.p)
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "first anchor must name a paragraph", firstAnchorId);
        if (secondEl.Name != W.p)
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "second anchor must name a paragraph", secondAnchorId);

        if (!ReferenceEquals(firstEl.NextNode, secondEl))
            return EditResult.Fail(EditErrorCode.AnchorsNotAdjacent,
                "MergeParagraphs requires second anchor to be the immediate next sibling of first");
        bool tracked = _trackedChanges == TrackedChangeMode.RenderInline;
        if (tracked && RefuseTrackedMerge(firstEl, secondEl, firstAnchorId) is { } trackedRefusal)
            return trackedRefusal;

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var stamp = tracked ? NewDistinctRevisionStamp() : default;
            // Insert a single-space separator if both sides end/start with non-whitespace.
            // Sentences from two paragraphs should not jam into one another.
            var firstTail = ParagraphText(firstEl);
            var secondHead = ParagraphText(secondEl);
            if (firstTail.Length > 0 && secondHead.Length > 0
                && !char.IsWhiteSpace(firstTail[^1])
                && !char.IsWhiteSpace(secondHead[0]))
            {
                var separator = new XElement(W.r,
                    new XElement(W.t,
                        new XAttribute(XNamespace.Xml + "space", "preserve"), " "));
                firstEl.Add(tracked ? CreateRevisionEnvelope(W.ins, stamp, separator) : separator);
            }

            if (tracked)
            {
                // Word's tracked join: the first paragraph's mark is deleted, and both paragraphs
                // stay until the deletion is accepted.
                TrackMergedProperties(firstEl.Element(W.pPr), secondEl, stamp);
                MarkParagraphMark(firstEl, W.del, stamp.Author, stamp.Date);
                InvalidateProjectionCache();
                // The second paragraph took the first's properties, so its kind can change
                // (p → h, p → li): report the anchors as the fresh index has them.
                return new EditResult
                {
                    Success = true,
                    Modified = new[]
                    {
                        AnchorForUnid(firstTarget.Unid, firstTarget.PartUri) ?? firstTarget.Anchor,
                        AnchorForUnid(secondTarget.Unid, secondTarget.PartUri) ?? secondTarget.Anchor,
                    },
                    Patch = PatchFor(firstTarget),
                };
            }

            // The second mark survives the join, so its section boundary survives too.
            // Removing the first mark removes its boundary even if the second has none.
            firstEl.Element(W.pPr)?.Elements(W.sectPr).Remove();
            if (secondEl.Element(W.pPr)?.Element(W.sectPr) is { } section)
            {
                section.Remove();
                var pPr = firstEl.Element(W.pPr);
                if (pPr is null)
                {
                    pPr = new XElement(W.pPr);
                    firstEl.AddFirst(pPr);
                }
                SetPPrChildInOrder(pPr, section);
            }

            // Move every paragraph-level child from secondEl into firstEl in document
            // order — runs, hyperlinks, sdts, fldSimples, bookmarkStart/End, comment
            // range markers, etc. The old implementation only moved direct <w:r>
            // children which silently discarded everything else.
            foreach (var child in secondEl.Elements().ToList())
            {
                if (child.Name == W.pPr) continue; // other properties come from the first
                child.Remove();
                firstEl.Add(child);
            }
            secondEl.Remove();
            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Modified = new[] { firstTarget.Anchor },
                Removed = new[] { secondTarget.Anchor },
                Patch = PatchFor(firstTarget),
            };
        }
        catch (Exception ex)
        {
            return FailInternal(ex);
        }
    }
}
