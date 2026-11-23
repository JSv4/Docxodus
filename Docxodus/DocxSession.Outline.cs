// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

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
    /// counts as a heading by its outline level, not by its name. <c>w:val="9"</c> is body text.
    /// </summary>
    public IReadOnlyList<OutlineEntry> ListOutline()
    {
        ThrowIfDisposed();
        _ = AnchorIndex();
        var body = _doc!.MainDocumentPart?.GetXDocument().Root?.Element(W.body);
        if (body is null) return Array.Empty<OutlineEntry>();
        var styles = _doc.MainDocumentPart!.StyleDefinitionsPart?.GetXDocument().Root;
        var byId = styles?.Elements(W.style)
            .Where(s => (string?)s.Attribute(W.type) == "paragraph" && s.Attribute(W.styleId) is not null)
            .GroupBy(s => (string)s.Attribute(W.styleId)!)
            .ToDictionary(g => g.Key, g => g.First()) ?? new Dictionary<string, XElement>();
        var defaultStyle = styles?.Elements(W.style).FirstOrDefault(s =>
            (string?)s.Attribute(W.type) == "paragraph" && (string?)s.Attribute(W._default) is "1" or "true");
        var levelOfStyle = new Dictionary<string, int?>();

        int? StyleLevel(string styleId)
        {
            if (levelOfStyle.TryGetValue(styleId, out var known)) return known;
            int? level = null;
            var seen = new HashSet<string>();
            for (var id = styleId; id is not null && seen.Add(id) && byId.TryGetValue(id, out var style);
                 id = (string?)style.Element(W.basedOn)?.Attribute(W.val))
            {
                if (OutlineLevelOf(style.Element(W.pPr)) is { } own) { level = own; break; }
            }
            return levelOfStyle[styleId] = level;
        }

        // The top-level blocks in order (a block content control's are among them), each heading's
        // position in them, then each heading's part: up to the next heading at its level or above.
        var blocks = body.Descendants()
            .Where(e => (e.Name == W.p || e.Name == W.tbl)
                && !e.Ancestors().Any(a => a.Name == W.tbl || a.Name == W.txbxContent || a.Name == MC.Fallback))
            .ToList();
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
            if (AnchorForElement(p) is not { } anchor) continue;
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
                if (AnchorForElement(blocks[j]) is { } a) { last = a.Id; break; }
            result.Add(new OutlineEntry(anchor.Id, level, VisibleText(p),
                (string?)p.Element(W.pPr)?.Element(W.pStyle)?.Attribute(W.val), last));
        }
        return result;
    }

    /// <summary>A <c>w:pPr</c>'s own <c>w:outlineLvl</c> (0-9), or null when it sets none.</summary>
    private static int? OutlineLevelOf(XElement? pPr) =>
        int.TryParse((string?)pPr?.Element(W.outlineLvl)?.Attribute(W.val), out var level) ? level : null;

    /// <summary>
    /// Move the run of top-level body blocks from <paramref name="firstAnchorId"/> to
    /// <paramref name="lastAnchorId"/> (inclusive, in document order; paragraphs and tables) before
    /// or after <paramref name="targetAnchorId"/> — a heading and everything under it, as Word's
    /// Navigation Pane drags it. One undo step. The same guards as <see cref="MoveBlock"/>: refused
    /// when the run or the way it travels crosses a section break, when it would cut or invert a
    /// cross-block range (a bookmark or comment running into the run from outside), when the target
    /// lies inside the run, and under tracked changes (a tracked multi-block move is not supported).
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
        if (first.Parent is not { } parent || !ReferenceEquals(last.Parent, parent) || !ReferenceEquals(target.Parent, parent))
            return EditResult.Fail(EditErrorCode.InvalidPosition,
                "MoveBlockRange's blocks and target must share a direct XML parent", firstAnchorId);
        if (!IsEditorBodyBlock(first, parent) || !IsEditorBodyBlock(last, parent) || !IsEditorBodyBlock(target, parent))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                "MoveBlockRange only supports top-level body blocks", firstAnchorId);

        var ctx = BuildBlockMoveContext(parent);
        if (!ctx.Index.TryGetValue(first, out var s0) || !ctx.Index.TryGetValue(last, out var s1)
            || !ctx.Index.TryGetValue(target, out var t))
            return EditResult.Fail(EditErrorCode.InvalidPosition, "a block is not a body render unit", firstAnchorId);
        if (s1 < s0)
            return EditResult.Fail(EditErrorCode.InvalidPosition, "the last block comes before the first", lastAnchorId);
        if (t >= s0 && t <= s1)
            return EditResult.Fail(EditErrorCode.InvalidPosition, "the target lies inside the blocks to move", targetAnchorId);
        // Already there: a no-op that takes no undo step.
        if ((pos == Position.Before && t == s1 + 1) || (pos == Position.After && t == s0 - 1))
            return new EditResult { Success = true };

        var run = ctx.Blocks.Skip(s0).Take(s1 - s0 + 1).ToList();
        foreach (var block in run)
            if (MoveSourceRejection(block) is { } why)
                return EditResult.Fail(EditErrorCode.InvalidPosition, why, firstAnchorId);
        int lo = Math.Min(s0, t), hi = Math.Max(s1, t);
        if (ctx.SectionBreaksBefore[hi + 1] - ctx.SectionBreaksBefore[lo] > 0)
            return EditResult.Fail(EditErrorCode.InvalidPosition, "cannot move blocks across a section-break paragraph", firstAnchorId);
        if (ctx.UniversalRejection is { } universal)
            return EditResult.Fail(EditErrorCode.InvalidPosition, universal, firstAnchorId);

        // Where each block lands once the run is lifted out and put back at the target: a
        // cross-block range must keep its two ends in order and the same distance apart.
        int n = s1 - s0 + 1;
        int targetAfterRemoval = t > s1 ? t - n : t;
        int insertAt = pos == Position.Before ? targetAfterRemoval : targetAfterRemoval + 1;
        int Reordered(int i)
        {
            if (i >= s0 && i <= s1) return insertAt + (i - s0);
            int afterRemoval = i > s1 ? i - n : i;
            return afterRemoval >= insertAt ? afterRemoval + n : afterRemoval;
        }
        foreach (var (start, end, label) in ctx.CrossBlockRanges)
        {
            int a = Reordered(start), b = Reordered(end);
            if (b < a || b - a != end - start)
                return EditResult.Fail(EditErrorCode.InvalidPosition,
                    $"move would change or invert a cross-block {label} range", firstAnchorId);
        }

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            // The blocks' own siblings between them (a run-level marker left among blocks, say)
            // travel with them: everything from the first block to the last.
            var nodes = new List<XNode>();
            for (XNode? node = first; node is not null; node = node.NextNode)
            {
                nodes.Add(node);
                if (ReferenceEquals(node, last)) break;
            }
            foreach (var node in nodes) node.Remove();
            if (pos == Position.Before) target.AddBeforeSelf(nodes);
            else target.AddAfterSelf(nodes);
            InvalidateProjectionCache();
            Anchor[] moved = run.Select(AnchorForElement).OfType<Anchor>().ToArray();
            return new EditResult { Success = true, Modified = moved.Length > 0 ? moved : new[] { firstTarget.Anchor } };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, firstAnchorId);
        }
    }
}
