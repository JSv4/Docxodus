// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Docxodus.Internal;

namespace Docxodus;

/// <summary>One heading of the document's outline (see <see cref="DocxSession.ListOutline"/>).</summary>
/// <param name="AnchorId">The heading paragraph's anchor.</param>
/// <param name="Level">Its outline level, 1-9 (Word's Level 1 is <c>w:outlineLvl w:val="0"</c>).</param>
/// <param name="Text">Its visible text.</param>
/// <param name="StyleId">Its paragraph style, or null when it has none.</param>
/// <param name="LastAnchorId">The last top-level block (paragraph or table) of the heading's part
/// of the document: everything up to the next heading of its level or a higher one (the end of
/// the body for the last). What <see cref="DocxSession.MoveBlockRange"/> moves with the heading.</param>
public sealed record OutlineEntry(string AnchorId, int Level, string Text, string? StyleId, string LastAnchorId);

public sealed partial class DocxSession
{
    /// <summary>
    /// The body's headings in document order, as Word's Navigation Pane shows them: every top-level
    /// body paragraph (not one in a table cell or a text box) whose outline level is 1-9. The level
    /// is resolved as Word resolves it: the paragraph's own <c>w:outlineLvl</c>, else its style's,
    /// following <c>w:basedOn</c> — so a localized heading style ("Título 1", id <c>Ttulo1</c>)
    /// counts as a heading by its outline level, not by its name. <c>w:val="9"</c> is body text. A
    /// built-in heading style that writes no level (some producers leave it out) has its own.
    /// </summary>
    public IReadOnlyList<OutlineEntry> ListOutline()
    {
        ThrowIfDisposed();
        var anchors = BodyAnchors();
        var body = _doc!.MainDocumentPart?.GetXDocument().Root?.Element(W.body);
        if (body is null) return Array.Empty<OutlineEntry>();
        var styles = _doc.MainDocumentPart!.StyleDefinitionsPart?.GetXDocument().Root;
        var byId = styles?.Elements(W.style)
            .Where(s => (string?)s.Attribute(W.type) == "paragraph" && s.Attribute(W.styleId) is not null)
            .GroupBy(s => (string)s.Attribute(W.styleId)!)
            .ToDictionary(g => g.Key, g => g.First()) ?? new Dictionary<string, XElement>();
        var defaultStyle = styles?.Elements(W.style).FirstOrDefault(s =>
            (string?)s.Attribute(W.type) == "paragraph" && WordprocessingMLUtil.OnOffValue((string?)s.Attribute(W._default)) == true);
        var levelOfStyle = new Dictionary<string, int?>();

        int? StyleLevel(string styleId)
        {
            if (levelOfStyle.TryGetValue(styleId, out var known)) return known;
            int? level = null;
            var seen = new HashSet<string>();
            for (var id = styleId; id is not null && seen.Add(id) && byId.TryGetValue(id, out var style);
                 id = (string?)style.Element(W.basedOn)?.Attribute(W.val))
            {
                if ((OutlineLevelOf(style.Element(W.pPr)) ?? BuiltInHeadingLevelOf(style)) is { } own) { level = own; break; }
            }
            return levelOfStyle[styleId] = level;
        }

        // The top-level blocks in order (a block content control's are among them), each heading's
        // position in them, then each heading's part: up to the next heading at its level or above.
        var blocks = TopLevelBlocks(body).ToList();
        var headings = new List<(int Index, Anchor Anchor, int Level, XElement P)>();
        for (int i = 0; i < blocks.Count; i++)
        {
            var p = blocks[i];
            if (p.Name != W.p) continue;
            var pPr = p.Element(W.pPr);
            var styleId = (string?)pPr?.Element(W.pStyle)?.Attribute(W.val)
                ?? (string?)defaultStyle?.Attribute(W.styleId);
            var level = OutlineLevelOf(pPr) ?? (styleId is null ? null : StyleLevel(styleId));
            if (level is not (>= 0 and <= 8)) continue;
            if (AnchorIn(anchors, p) is not { } anchor) continue;
            headings.Add((i, anchor, level.Value + 1, p));
        }
        var result = new List<OutlineEntry>(headings.Count);
        for (int h = 0; h < headings.Count; h++)
        {
            var (index, anchor, level, p) = headings[h];
            int end = blocks.Count - 1;
            for (int k = h + 1; k < headings.Count; k++)
                if (headings[k].Level <= level) { end = headings[k].Index - 1; break; }
            // (the last block with an anchor: a section-only paragraph has one too)
            string last = anchor.Id;
            for (int j = end; j > index; j--)
                if (AnchorIn(anchors, blocks[j]) is { } a) { last = a.Id; break; }
            result.Add(new OutlineEntry(anchor.Id, level, ParagraphText(p),
                (string?)p.Element(W.pPr)?.Element(W.pStyle)?.Attribute(W.val), last));
        }
        return result;
    }

    /// <summary>A <c>w:pPr</c>'s own <c>w:outlineLvl</c> (0-9), or null when it sets none.</summary>
    private static int? OutlineLevelOf(XElement? pPr) =>
        int.TryParse((string?)pPr?.Element(W.outlineLvl)?.Attribute(W.val), out var level) ? level : null;

    /// <summary>The outline level (0-8) of a built-in heading style, which its name fixes
    /// (<c>heading 1</c> to <c>heading 9</c> in every language), or null for any other style.</summary>
    private static int? BuiltInHeadingLevelOf(XElement style) =>
        (string?)style.Element(W.name)?.Attribute(W.val) is { Length: 9 } name
        && name.StartsWith("heading ", StringComparison.OrdinalIgnoreCase) && name[8] is >= '1' and <= '9'
            ? name[8] - '1' : null;

    /// <summary>The paragraphs and tables in <paramref name="root"/> (itself included) that are in no
    /// table cell or text box, in document order.</summary>
    private static IEnumerable<XElement> TopLevelBlocks(XElement root) =>
        root.DescendantsAndSelf().Where(e => (e.Name == W.p || e.Name == W.tbl)
            && !e.Ancestors().Any(a => a.Name == W.tbl || a.Name == W.txbxContent) && WmlToMarkdownConverter.IsSelectedMarkup(e));

    /// <summary>
    /// Move the run of top-level body blocks from <paramref name="firstAnchorId"/> to
    /// <paramref name="lastAnchorId"/> (inclusive, in document order; paragraphs and tables) before
    /// or after <paramref name="targetAnchorId"/> — a heading and everything under it, as Word's
    /// Navigation Pane drags it. One undo step. A block content control moves whole or not at all:
    /// a run that reaches into one must cover it, and a target inside one stands for the place just
    /// outside it. Refused when the run or the way it travels crosses a section break, when the move
    /// would cut, invert or change what a bookmark, comment, permission or field covers, when the
    /// target lies inside the run, and under tracked changes (a tracked multi-block move is not
    /// supported).
    /// </summary>
    public EditResult MoveBlockRange(string firstAnchorId, string lastAnchorId, string targetAnchorId, Position pos)
    {
        if (MutationRefusal() is { } refusal) return EditResult.Fail(refusal);
        var firstTarget = FindAnchor(firstAnchorId);
        var lastTarget = FindAnchor(lastAnchorId);
        var targetTarget = FindAnchor(targetAnchorId);
        if (firstTarget is null || lastTarget is null || targetTarget is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "first, last or target anchor not found", firstAnchorId);
        foreach (var given in new[] { firstTarget, lastTarget, targetTarget })
            if (given.Anchor.Kind is not ("p" or "h" or "li" or "tbl"))
                return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                    $"MoveBlockRange requires paragraph/heading/list/table anchors; got kind={given.Anchor.Kind}", given.Anchor.Id);
        if (_trackedChanges == TrackedChangeMode.RenderInline)
            return TrackedStructureUnsupported("MoveBlockRange", firstAnchorId);

        var first = firstTarget.Resolve(_doc!);
        var last = lastTarget.Resolve(_doc!);
        var target = targetTarget.Resolve(_doc!);
        if (first is null || last is null || target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "an anchor resolved null", firstAnchorId);
        // The closest body or block content control that holds all three: its children are what moves.
        var body = first.Ancestors(W.body).FirstOrDefault();
        var container = first.Ancestors().FirstOrDefault(a => (a.Name == W.body || a.Name == W.sdtContent)
            && last.Ancestors().Contains(a) && target.Ancestors().Contains(a));
        if (body is null || container is null
            || new[] { first, last, target }.Any(b => !TopLevelBlocks(b).Contains(b)))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                "MoveBlockRange only supports top-level body blocks", firstAnchorId);

        XElement Unit(XElement block) => block.AncestorsAndSelf().First(e => ReferenceEquals(e.Parent, container));
        bool OpensUnit(XElement block) => ReferenceEquals(TopLevelBlocks(Unit(block)).First(), block);
        bool ClosesUnit(XElement block) => ReferenceEquals(TopLevelBlocks(Unit(block)).Last(), block);
        if (!OpensUnit(first) || !ClosesUnit(last) || !(pos == Position.Before ? OpensUnit(target) : ClosesUnit(target)))
            return EditResult.Fail(EditErrorCode.InvalidPosition,
                "the move would split a content control, or put blocks inside one", firstAnchorId);

        var children = container.Elements().ToList();
        var index = new Dictionary<XElement, int>(children.Count);
        for (int i = 0; i < children.Count; i++) index[children[i]] = i;
        int s0 = index[Unit(first)], s1 = index[Unit(last)], t = index[Unit(target)];
        if (s1 < s0)
            return EditResult.Fail(EditErrorCode.InvalidPosition, "the last block comes before the first", lastAnchorId);
        if (t >= s0 && t <= s1)
            return EditResult.Fail(EditErrorCode.InvalidPosition, "the target lies inside the blocks to move", targetAnchorId);

        // Every range and field, as the children its two ends are in (or are); one end outside the
        // container is past its first or last child.
        int At(XElement marker) =>
            marker.AncestorsAndSelf().FirstOrDefault(e => ReferenceEquals(e.Parent, container)) is { } child ? index[child]
            : XNode.DocumentOrderComparer.Compare(marker, container) < 0 ? -1 : children.Count;
        var ranges = CrossBlockRangesAndFields(body).Select(r => (Start: At(r.Start), End: At(r.End), r.Label)).ToList();
        // A marker right beside the run (Word writes the end of a bookmark over whole paragraphs there)
        // travels with it when its other end does.
        bool Travels(int marker) => ranges.Any(r =>
            (r.Start == marker && r.End >= s0 && r.End <= s1) || (r.End == marker && r.Start >= s0 && r.Start <= s1));
        for (bool grew = true; grew;)
        {
            grew = false;
            if (s1 + 1 < children.Count && IsBlockLevelMarker(children[s1 + 1]) && Travels(s1 + 1)) { s1++; grew = true; }
            if (s0 > 0 && IsBlockLevelMarker(children[s0 - 1]) && Travels(s0 - 1)) { s0--; grew = true; }
        }

        // Where it can land: beside the target, on either side of the markers between it and the next block.
        int n = s1 - s0 + 1;
        var rest = children.Where((_, i) => i < s0 || i > s1).ToList();
        int at = t > s1 ? t - n : t;
        var slots = new List<int> { pos == Position.Before ? at : at + 1 };
        if (pos == Position.Before)
            while (slots[^1] > 0 && IsBlockLevelMarker(rest[slots[^1] - 1])) slots.Add(slots[^1] - 1);
        else
            while (slots[^1] < rest.Count && IsBlockLevelMarker(rest[slots[^1]])) slots.Add(slots[^1] + 1);
        // Already there: a no-op that takes no undo step.
        if (slots.Contains(s0)) return new EditResult { Success = true };

        var run = children.GetRange(s0, n);
        foreach (var block in run)
            if (MoveSourceRejection(block) is { } why)
                return EditResult.Fail(EditErrorCode.InvalidPosition, why, firstAnchorId);
        if (children.Skip(Math.Min(s0, t)).Take(Math.Max(s1, t) - Math.Min(s0, t) + 1)
            .Any(c => StoryParagraphs(c).Any(p => p.Element(W.pPr)?.Element(W.sectPr) is not null)))
            return EditResult.Fail(EditErrorCode.InvalidPosition, "cannot move blocks across a section-break paragraph", firstAnchorId);

        // A range or field must keep its two ends in order and the same distance apart.
        string? Broken(int slot)
        {
            int Reordered(int i)
            {
                if (i >= s0 && i <= s1) return slot + (i - s0);
                int r = i > s1 ? i - n : i;
                return r >= slot ? r + n : r;
            }
            foreach (var (start, end, label) in ranges)
                if (end < start || Reordered(end) - Reordered(start) != end - start) return label;
            return null;
        }
        int chosen = slots.FindIndex(slot => Broken(slot) is null);
        if (chosen < 0)
            return EditResult.Fail(EditErrorCode.InvalidPosition,
                $"move would change or invert a cross-block {Broken(slots[0])} range", firstAnchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            // The blocks' own siblings between them travel with them: everything from the first to the last.
            var nodes = new List<XNode>();
            for (XNode? node = run[0]; node is not null; node = node.NextNode)
            {
                nodes.Add(node);
                if (ReferenceEquals(node, run[^1])) break;
            }
            var (gapBefore, gapAfter) = (RevisionOps.SiblingBlock(run[0], forward: false), RevisionOps.SiblingBlock(run[^1], forward: true));
            foreach (var node in nodes) node.Remove();
            int slot = slots[chosen];
            if (slot < rest.Count) rest[slot].AddBeforeSelf(nodes);
            else rest[^1].AddAfterSelf(nodes);
            // Where the blocks came from, and both edges of where they went: a paragraph must stand between
            // two tables, and after one that would end the body (as InsertTable keeps it).
            var separators = new List<XElement>();
            // (the run's own first and last block: a marker that travels with it may be outside them)
            XElement firstBlock = run.First(b => !IsBlockLevelMarker(b)), lastBlock = run.Last(b => !IsBlockLevelMarker(b));
            foreach (var (left, right) in new[]
                {
                    (gapBefore, gapAfter),
                    (RevisionOps.SiblingBlock(firstBlock, forward: false), firstBlock),
                    (lastBlock, RevisionOps.SiblingBlock(lastBlock, forward: true)),
                })
            {
                if (left?.Name != W.tbl || !ReferenceEquals(RevisionOps.SiblingBlock(left, forward: true), right)
                    || !(right?.Name == W.tbl || right?.Name == W.sectPr || (right is null && container.Name == W.body))) continue;
                var separator = new XElement(W.p);
                UnidHelper.AssignToSelfAndDescendants(separator);
                // After the ends of the ranges that cover the table, so that none of them covers it.
                var after = left;
                while (after.ElementsAfterSelf().FirstOrDefault() is { } next && CrossBlockRangePairs.Any(pair => next.Name == pair.End)) after = next;
                after.AddAfterSelf(separator);
                separators.Add(separator);
            }
            InvalidateProjectionCache();
            var anchors = BodyAnchors();
            return new EditResult
            {
                Success = true,
                Created = separators.Select(p => AnchorIn(anchors, p)).OfType<Anchor>().ToArray(),
                Modified = run.SelectMany(TopLevelBlocks).Select(b => AnchorIn(anchors, b)).OfType<Anchor>().ToArray(),
                Patch = PatchFor(firstTarget),
            };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, firstAnchorId);
        }
    }

    /// <summary>The main document's anchors by unid, for an operation that looks up many blocks: each
    /// <see cref="AnchorForElement"/> is a scan of the whole index.</summary>
    private Dictionary<string, Anchor> BodyAnchors()
    {
        var part = _doc!.MainDocumentPart!.Uri.ToString();
        var anchors = new Dictionary<string, Anchor>(StringComparer.Ordinal);
        foreach (var target in AnchorIndex().Values)
            if (target.PartUri == part) anchors.TryAdd(target.Unid, target.Anchor);
        return anchors;
    }

    private static Anchor? AnchorIn(Dictionary<string, Anchor> anchors, XElement block) =>
        (string?)block.Attribute(PtOpenXml.Unid) is { } unid && anchors.TryGetValue(unid, out var anchor) ? anchor : null;

    /// <summary>The story's own paragraphs in <paramref name="root"/>, table cells included: a text box's
    /// are another story, whose section properties break no section of this one.</summary>
    private static IEnumerable<XElement> StoryParagraphs(XElement root) =>
        root.DescendantsAndSelf(W.p).Where(p => !p.Ancestors().Any(a => a.Name == W.txbxContent));

    /// <summary>A range marker (or a proofing mark) that is itself a child of the body or of a block
    /// content control, so not a block.</summary>
    private static bool IsBlockLevelMarker(XElement child) => child.Name == W.proofErr
        || CrossBlockRangePairs.Any(pair => child.Name == pair.Start || child.Name == pair.End);

    /// <summary>The two ends of every paired range marker and every complex field in
    /// <paramref name="root"/>.</summary>
    private static IEnumerable<(XElement Start, XElement End, string Label)> CrossBlockRangesAndFields(XElement root)
    {
        var starts = new Dictionary<(XName Name, string Id), XElement>();
        var fields = new Stack<XElement>();
        foreach (var e in root.Descendants())
        {
            if (e.Name == W.fldChar)
            {
                var type = (string?)e.Attribute(W.fldCharType);
                if (type == "begin") fields.Push(e);
                else if (type == "end" && fields.Count > 0) yield return (fields.Pop(), e, "field");
                continue;
            }
            foreach (var (startName, endName, _) in CrossBlockRangePairs)
                if (e.Name == startName) starts.TryAdd((endName, (string?)e.Attribute(W.id) ?? ""), e);
        }
        foreach (var e in root.Descendants())
            foreach (var (_, endName, label) in CrossBlockRangePairs)
                if (e.Name == endName && starts.TryGetValue((endName, (string?)e.Attribute(W.id) ?? ""), out var start))
                    yield return (start, e, label);
    }
}
