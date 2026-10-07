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
    /// paragraph move or a Word-native deleted/inserted table pair.
    /// </summary>
    public EditResult MoveBlock(string sourceAnchorId, string targetAnchorId, Position pos)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");

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

        // A source already in the requested slot is a true no-op: do not consume undo.
        if ((pos == Position.Before && ReferenceEquals(source.NextNode, target)) ||
            (pos == Position.After && ReferenceEquals(target.NextNode, source)))
            return new EditResult { Success = true };

        if (TrackedBookmarkMoveRejection(source) is { } bookmarkRejection)
            return EditResult.Fail(EditErrorCode.UnsupportedInlineBoundary, bookmarkRejection, sourceAnchorId);

        if (MoveSourceRejection(source) is { } sourceRejection)
            return EditResult.Fail(EditErrorCode.InvalidPosition, sourceRejection, sourceAnchorId);
        if (BlockMoveSafetyError(BuildBlockMoveContext(parent), source, target, pos) is { } safetyError)
            return EditResult.Fail(EditErrorCode.InvalidPosition, safetyError, sourceAnchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            if (_trackedChanges != TrackedChangeMode.RenderInline)
            {
                source.Remove();
                if (pos == Position.Before) target.AddBeforeSelf(source);
                else target.AddAfterSelf(source);

                InvalidateProjectionCache();
                return new EditResult
                {
                    Success = true,
                    Modified = new[] { sourceTarget.Anchor },
                    Patch = PatchFor(sourceTarget),
                };
            }

            var destination = new XElement(source);
            foreach (var el in destination.DescendantsAndSelf())
                el.Attributes(PtOpenXml.Unid).Remove();

            // Both copies are live while the revision is pending, so the shared comment ids have to
            // be split. The move SOURCE takes the clones (see CloneCommentsForMoveSource), leaving
            // the destination wired to the original comments and their threads.
            if (_doc!.MainDocumentPart is { } commentHost)
                Internal.CommentOps.CloneCommentsForMoveSource(commentHost, source);

            var author = _revisionAuthor ?? "docxodus";
            var date = RevisionDateNow();
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
            if (pos == Position.Before) target.AddBeforeSelf(destination);
            else target.AddAfterSelf(destination);

            var destinationUnid = (string)destination.Attribute(PtOpenXml.Unid)!;
            var destinationAnchor = new Anchor(
                $"{sourceTarget.Anchor.Kind}:{sourceTarget.Anchor.Scope}:{destinationUnid}",
                sourceTarget.Anchor.Kind, sourceTarget.Anchor.Scope, destinationUnid);

            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Modified = new[] { sourceTarget.Anchor },
                Created = new[] { destinationAnchor },
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
        if (!IsEditorBodyBlock(source, parent) || MoveSourceRejection(source) is not null)
            return empty;

        // One context for the whole sweep: each candidate then costs index arithmetic over the
        // container's cross-block ranges instead of its own pass over the story.
        var context = BuildBlockMoveContext(parent);
        if (context.UniversalRejection is not null) return empty;

        var targets = new List<MoveTarget>();
        foreach (var candidate in context.Blocks)
        {
            if (ReferenceEquals(candidate, source) || !IsEditorBodyBlock(candidate, parent))
                continue;
            var unid = (string?)candidate.Attribute(PtOpenXml.Unid);
            if (unid is null) continue;
            var kind = candidate.Name == W.tbl ? "tbl" : WmlToMarkdownConverter.KindFor(candidate);
            if (kind is null) continue;

            bool before = BlockMoveSafetyError(context, source, candidate, Position.Before) is null;
            bool after = BlockMoveSafetyError(context, source, candidate, Position.After) is null;
            if (before || after)
                targets.Add(new MoveTarget($"{kind}:{sourceTarget.Anchor.Scope}:{unid}", before, after));
        }
        return targets;
    }

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
    /// The move-independent facts about one block container: the body render units in document
    /// order, which of them own a section break, and where each cross-block comment/bookmark/
    /// permission/native-move range starts and ends.
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
        /// <summary>Body render units in document order.</summary>
        public required IReadOnlyList<XElement> Blocks { get; init; }

        /// <summary>Position of each block in <see cref="Blocks"/>, by reference.</summary>
        public required Dictionary<XElement, int> Index { get; init; }

        /// <summary>Number of section-break-owning blocks in <c>Blocks[0..i)</c>, so
        /// "is there one between these two blocks" is a subtraction.</summary>
        public required int[] SectionBreaksBefore { get; init; }

        /// <summary>Every range whose start and end sit in DIFFERENT blocks, as the
        /// index pair the move has to leave intact.</summary>
        public required IReadOnlyList<(int Start, int End, string Label)> CrossBlockRanges { get; init; }

        /// <summary>Set when the container holds an inverted range (its end precedes its
        /// start): that rejects every move, so no candidate needs testing.</summary>
        public string? UniversalRejection { get; init; }
    }

    private static BlockMoveContext BuildBlockMoveContext(XElement parent)
    {
        var blocks = parent.Elements().Where(e => e.Name == W.p || e.Name == W.tbl).ToList();
        // XElement does not override Equals, so the default comparer IS reference identity —
        // the same relation the previous List.IndexOf lookups used.
        var index = new Dictionary<XElement, int>(blocks.Count);
        for (int i = 0; i < blocks.Count; i++) index[blocks[i]] = i;

        var sectionBreaksBefore = new int[blocks.Count + 1];
        for (int i = 0; i < blocks.Count; i++)
        {
            bool owns = blocks[i].Name == W.p &&
                blocks[i].Element(W.pPr)?.Element(W.sectPr) is not null;
            sectionBreaksBefore[i + 1] = sectionBreaksBefore[i] + (owns ? 1 : 0);
        }

        // One traversal collects every marker of interest; the five names are matched
        // against each element rather than walking the story once per name.
        var starts = new Dictionary<XName, Dictionary<string, XElement>>();
        var ends = new Dictionary<XName, List<XElement>>();
        foreach (var (startName, endName, _) in CrossBlockRangePairs)
        {
            starts[startName] = new Dictionary<string, XElement>();
            ends[endName] = new List<XElement>();
        }
        foreach (var el in parent.Descendants())
        {
            if (starts.TryGetValue(el.Name, out var byId))
            {
                // First start wins for a duplicated id, matching the previous grouping.
                var id = (string?)el.Attribute(W.id) ?? "";
                if (!byId.ContainsKey(id)) byId[id] = el;
            }
            else if (ends.TryGetValue(el.Name, out var list))
            {
                list.Add(el);
            }
        }

        var ranges = new List<(int, int, string)>();
        string? universalRejection = null;
        foreach (var (startName, endName, label) in CrossBlockRangePairs)
        {
            var byId = starts[startName];
            foreach (var end in ends[endName])
            {
                var id = (string?)end.Attribute(W.id) ?? "";
                if (!byId.TryGetValue(id, out var start)) continue;
                var startBlock = TopLevelOwner(start, parent);
                var endBlock = TopLevelOwner(end, parent);
                if (startBlock is null || endBlock is null || ReferenceEquals(startBlock, endBlock))
                    continue;
                if (!index.TryGetValue(startBlock, out var a) || !index.TryGetValue(endBlock, out var b))
                    continue;
                if (b < a)
                {
                    // An inverted range was `before is null` for every candidate before.
                    universalRejection ??= $"move would change or invert a cross-block {label} range";
                    continue;
                }
                ranges.Add((a, b, label));
            }
        }

        return new BlockMoveContext
        {
            Blocks = blocks,
            Index = index,
            SectionBreaksBefore = sectionBreaksBefore,
            CrossBlockRanges = ranges,
            UniversalRejection = universalRejection,
        };
    }

    /// <summary>Reject a move when it would change the membership/order of a cross-block
    /// comment, bookmark, permission, or native-move range, or cross a section break.</summary>
    /// <remarks>
    /// A move relocates exactly ONE element, so the reordered document order is a function of
    /// three indices and needs no second list. A range survives iff its two endpoints still
    /// bound a window of the same width: the elements strictly between them can only change by
    /// the source entering or leaving, so equal width and equal membership are the same
    /// condition — including when the source IS an endpoint, where any relocation but the
    /// identity one changes the width.
    /// </remarks>
    private static string? BlockMoveSafetyError(
        BlockMoveContext ctx, XElement source, XElement target, Position pos)
    {
        if (!ctx.Index.TryGetValue(source, out var sourceIndex) ||
            !ctx.Index.TryGetValue(target, out var targetIndex))
            return "source or target is not a body render unit";

        int lo = Math.Min(sourceIndex, targetIndex);
        int hi = Math.Max(sourceIndex, targetIndex);
        if (ctx.SectionBreaksBefore[hi + 1] - ctx.SectionBreaksBefore[lo] > 0)
            return "cannot move a block across a section-break paragraph";

        // Checked after the section break, matching the order the messages were produced in
        // when the range scan ran per candidate.
        if (ctx.UniversalRejection is { } universal) return universal;

        // Where the source lands once it has been lifted out of the sequence.
        int targetAfterRemoval = targetIndex > sourceIndex ? targetIndex - 1 : targetIndex;
        int insertAt = pos == Position.Before ? targetAfterRemoval : targetAfterRemoval + 1;

        int Reordered(int i)
        {
            if (i == sourceIndex) return insertAt;
            int afterRemoval = i > sourceIndex ? i - 1 : i;
            return afterRemoval >= insertAt ? afterRemoval + 1 : afterRemoval;
        }

        foreach (var (start, end, label) in ctx.CrossBlockRanges)
        {
            int a = Reordered(start);
            int b = Reordered(end);
            if (b < a || b - a != end - start)
                return $"move would change or invert a cross-block {label} range";
        }
        return null;
    }

    private static XElement? TopLevelOwner(XElement marker, XElement parent) =>
        marker.Ancestors().FirstOrDefault(e => ReferenceEquals(e.Parent, parent) &&
            (e.Name == W.p || e.Name == W.tbl));

    /// <summary>
    /// Wrap the live runs of <paramref name="paragraph"/> in a
    /// <paramref name="wrapperName"/> revision envelope and optionally mark the paragraph mark to match.
    /// </summary>
    /// <remarks>
    /// Shared by whole-block deletion, text replacement, paragraph/table moves and paragraph/row insertion.
    /// A DELETING wrapper also converts
    /// <c>w:t</c>→<c>w:delText</c> (and <c>w:instrText</c>→<c>w:delInstrText</c>), which is what
    /// Word writes, what <c>IrMarkupRenderer.ConvertTextToDelText</c> produces, and what
    /// <see cref="RevisionProcessor"/>'s reject path swaps back. The paragraph mark belongs
    /// to the same operation unless the caller retains it (a replacement or a final retained paragraph).
    /// </remarks>
    private XElement? MarkParagraphContentAndMark(
        XElement paragraph, XName wrapperName, string author, string date,
        bool preserveParagraphMark = false, bool retainInlineStructure = false)
    {
        bool deleting = wrapperName == W.del || wrapperName == W.moveFrom;

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
            if (deleting) ConvertTextToDeletedText(element);
        }

        var pPr = paragraph.Element(W.pPr);
        if (preserveParagraphMark) return pPr;
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
        if (!marker
            && element.Ancestors().FirstOrDefault(e => e.Name == W.ins || e.Name == W.moveTo) is { } insertion
            && insertion.Name == W.ins
            && string.Equals((string?)insertion.Attribute(W.author), stamp.Author, StringComparison.Ordinal))
        {
            RemoveFromInsertion(element, insertion);
            return;
        }
        var envelope = CreateRevisionEnvelope(W.del, stamp);
        element.ReplaceWith(envelope);
        envelope.Add(element);
        ConvertTextToDeletedText(element);
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
    }

    /// <summary>Convert a deleted run-level element's text to its deleted spelling in place,
    /// mirroring <c>IrMarkupRenderer.ConvertTextToDelText</c>. Text under a revision nested
    /// inside the element — a text box paragraph's own insertion, say — keeps its spelling: that
    /// revision owns it, and renaming it would leave <c>w:delText</c> that no reject of THIS
    /// deletion could ever restore.</summary>
    private static void ConvertTextToDeletedText(XElement runLevel)
    {
        foreach (var t in runLevel.DescendantsAndSelf(W.t).ToList())
            if (!UnderNestedRevision(t, runLevel)) t.Name = W.delText;
        foreach (var instr in runLevel.DescendantsAndSelf(W.instrText).ToList())
            if (!UnderNestedRevision(instr, runLevel)) instr.Name = W.delInstrText;
    }

    private static bool UnderNestedRevision(XElement text, XElement runLevel) =>
        text.Ancestors().TakeWhile(a => !ReferenceEquals(a, runLevel))
            .Any(a => RevisionWrapperNames.Contains(a.Name));

    private static bool IsOwnInsertion(XElement? insertion, string author) =>
        insertion is not null
        && string.Equals((string?)insertion.Attribute(W.author), author, StringComparison.Ordinal);

    /// <summary>Whether a tracked deletion of <paramref name="root"/> removes content outright
    /// rather than marking it: a run (other than a marker-only one) inside the session author's
    /// own pending insertion, which <see cref="DeleteInlineElementInPlace"/> un-inserts.</summary>
    private static bool HasOwnInsertedContent(XElement root, string author) =>
        root.Descendants(W.r).Any(run => !IsMarkerOnlyRun(run)
            && !run.Ancestors().Any(a => a.Name == W.del || a.Name == W.moveFrom)
            && run.Ancestors().FirstOrDefault(a => a.Name == W.ins || a.Name == W.moveTo) is { } insertion
            && insertion.Name == W.ins && IsOwnInsertion(insertion, author));

    /// <summary>Whether <paramref name="paragraph"/> is, in its entirety, the session author's
    /// own pending insertion: its pilcrow carries the author's <c>w:ins</c> and every child is
    /// the author's insertion (directly, or through a hyperlink, field or smart tag holding only
    /// the author's insertions), with no range marker whose partner may lie elsewhere. Such a
    /// paragraph is un-inserted outright by a tracked deletion, as Word does.</summary>
    private static bool IsWhollyOwnInsertion(XElement paragraph, string author) =>
        IsOwnInsertion(paragraph.Element(W.pPr)?.Element(W.rPr)?.Element(W.ins), author)
        && paragraph.Elements().Where(child => child.Name != W.pPr).All(child => IsOwnInsertedContent(child, author))
        && !paragraph.Descendants().Any(d => d.Name == W.bookmarkStart || d.Name == W.bookmarkEnd
            || d.Name == W.commentRangeStart || d.Name == W.commentRangeEnd
            || d.Name == W.permStart || d.Name == W.permEnd
            || d.Name == W.moveFromRangeStart || d.Name == W.moveFromRangeEnd
            || d.Name == W.moveToRangeStart || d.Name == W.moveToRangeEnd);

    private static bool IsOwnInsertedContent(XElement element, string author) =>
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
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
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

    public EditResult SplitParagraph(string anchorId, int characterOffset)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
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
        if (_trackedChanges == TrackedChangeMode.RenderInline)
            return TrackedStructureUnsupported("SplitParagraph", anchorId);

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
            }

            // Split any run that straddles the offset (descends into hyperlinks/sdts),
            // then split any container (hyperlink) that still straddles, then move all
            // inline children + markers at-or-past the offset to `second`.
            SplitRunsAtOffset(element, characterOffset);
            SplitInlineContainersAtOffset(element, characterOffset);
            MoveInlineChildrenAfter(element, characterOffset, second);

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
                        var rebuilt = new XElement(W.pPr,
                            new XElement(W.pStyle, new XAttribute(W.val, nextStyle)));
                        newPPr.ReplaceWith(rebuilt);
                        newPPr = rebuilt;
                    }
                }

                // Re-mint Unids on the new paragraph's property subtree so cloned property elements
                // (jc, ind, numPr, ...) don't carry the original's Unid onto a second element.
                foreach (var el in newPPr.DescendantsAndSelf())
                    el.Attributes(PtOpenXml.Unid).Remove();
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

    public EditResult MergeParagraphs(string firstAnchorId, string secondAnchorId)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
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

        if (!ReferenceEquals(firstEl.NextNode, secondEl))
            return EditResult.Fail(EditErrorCode.AnchorsNotAdjacent,
                "MergeParagraphs requires second anchor to be the immediate next sibling of first");
        if (_trackedChanges == TrackedChangeMode.RenderInline)
            return TrackedStructureUnsupported("MergeParagraphs", firstAnchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            // Insert a single-space separator if both sides end/start with non-whitespace.
            // Sentences from two paragraphs should not jam into one another.
            var firstTail = ParagraphText(firstEl);
            var secondHead = ParagraphText(secondEl);
            if (firstTail.Length > 0 && secondHead.Length > 0
                && !char.IsWhiteSpace(firstTail[^1])
                && !char.IsWhiteSpace(secondHead[0]))
            {
                firstEl.Add(new XElement(W.r,
                    new XElement(W.t,
                        new XAttribute(XNamespace.Xml + "space", "preserve"), " ")));
            }

            // Move every paragraph-level child from secondEl into firstEl in document
            // order — runs, hyperlinks, sdts, fldSimples, bookmarkStart/End, comment
            // range markers, etc. The old implementation only moved direct <w:r>
            // children which silently discarded everything else.
            foreach (var child in secondEl.Elements().ToList())
            {
                if (child.Name == W.pPr) continue; // second's pPr is dropped; first's wins
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
