// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Linq;
using System.Text;
using Docxodus.Ir;

namespace Docxodus.Ir.Diff;

/// <summary>
/// Pairs whole paragraphs that moved between a table cell and the body, or between two cells, as one
/// relocation (issue #887). The aligner detects moves within one block list: the body, or one cell. A
/// paragraph that crosses a table boundary therefore reaches the edit script as a <see cref="IrEditOpKind.DeleteBlock"/>
/// in one list and an <see cref="IrEditOpKind.InsertBlock"/> in another. This pass gives such a pair a shared
/// <see cref="IrEditOp.RelocationGroupId"/> and changes nothing else: the ops keep their kind, so applying the
/// script, and accepting or rejecting its markup, is exactly what it was. The renderers draw a tagged pair
/// as <c>w:moveFrom</c>/<c>w:moveTo</c>, and <c>GetRevisions</c> reports it as a Moved pair, when move
/// reporting is on.
/// <para><b>What pairs.</b> The two paragraphs' text must be identical once whitespace runs are collapsed,
/// and carry at least <see cref="IrDiffSettings.MoveMinimumTokenCount"/> words. An exact match leaves no
/// in-move edit for either surface to describe. Each paragraph pairs at most once, the first destination in
/// document order winning.</para>
/// <para><b>Where it looks.</b> The body and the cells of tables drawn cell by cell on both surfaces
/// (<see cref="IrTableDiffer.NeedsWholeTableFallback"/> false). A moved table is drawn whole, and its cells are
/// not looked into either. A paragraph the redline draws inside a cross-paragraph run
/// (<see cref="IrEditOpKind.CrossParagraphRunBlock"/>) is left out: the markup script has no standalone op
/// for it, and the data script is told to leave out the same paragraphs, so the two scripts pair the same
/// paragraphs. With <see cref="IrDiffSettings.PreserveInputRevisions"/> on, nothing pairs: the markup
/// renderer may draw a table whole to keep its input revisions, which the revision list cannot see.</para>
/// <para>Only the two-way comparison calls this; Consolidate's per-reviewer scripts never carry relocations.</para>
/// </summary>
internal static class IrRelocationPairer
{
    private sealed record Candidate(IrEditOp Op, int ListId, string Text);

    /// <param name="fusedAnchors">Body paragraph anchors the redline draws inside cross-paragraph runs, when
    /// this script is the data script of a comparison whose redline fuses runs; null otherwise.</param>
    public static IrEditScript Apply(
        IrEditScript script, IrDocument left, IrDocument right, IrDiffSettings settings,
        IReadOnlySet<string>? fusedAnchors = null)
    {
        if (settings.PreserveInputRevisions)
            return script;

        var deleted = new List<Candidate>();
        var inserted = new List<Candidate>();
        int nextList = 0;
        void Collect(IEnumerable<IrEditOp> ops)
        {
            int listId = nextList++;
            foreach (var op in ops)
            {
                if (op.Kind == IrEditOpKind.DeleteBlock && Candidate(op, op.LeftAnchor, left, listId) is { } d)
                    deleted.Add(d);
                else if (op.Kind == IrEditOpKind.InsertBlock && Candidate(op, op.RightAnchor, right, listId) is { } i)
                    inserted.Add(i);
                foreach (var cellOps in CellOpLists(op, settings))
                    Collect(cellOps);
            }
        }
        Candidate? Candidate(IrEditOp op, string? anchor, IrDocument doc, int listId)
        {
            if (anchor is null || fusedAnchors?.Contains(anchor) == true ||
                !doc.AnchorIndex.TryGetValue(anchor, out var block) || block is not IrParagraph paragraph)
                return null;
            var (text, words) = Normalize(IrDiffTokenizer.Tokenize(paragraph, settings));
            return words >= settings.MoveMinimumTokenCount ? new Candidate(op, listId, text) : null;
        }
        Collect(script.Operations);
        if (deleted.Count == 0 || inserted.Count == 0 || nextList == 1)
            return script;

        // Destinations by text, in document order; each source claims the first unused one in another list.
        var byText = inserted.GroupBy(i => i.Text, System.StringComparer.Ordinal)
            .ToDictionary(g => g.Key, g => g.ToList(), System.StringComparer.Ordinal);
        var pairs = new List<(Candidate Source, Candidate Destination)>();
        var used = new HashSet<Candidate>(ReferenceEqualityComparer.Instance);
        foreach (var source in deleted)
        {
            if (!byText.TryGetValue(source.Text, out var destinations))
                continue;
            var destination = destinations.FirstOrDefault(d => d.ListId != source.ListId && !used.Contains(d));
            if (destination is null)
                continue;
            used.Add(destination);
            pairs.Add((source, destination));
        }
        if (pairs.Count == 0)
            return script;

        // Ids continue above every move group and follow destination document order.
        int nextGroup = MaxMoveGroupId(script) + 1;
        var destinationOrder = new Dictionary<Candidate, int>(ReferenceEqualityComparer.Instance);
        for (int i = 0; i < inserted.Count; i++)
            destinationOrder[inserted[i]] = i;
        var ids = new Dictionary<IrEditOp, int>(ReferenceEqualityComparer.Instance);
        foreach (var (source, destination) in pairs.OrderBy(p => destinationOrder[p.Destination]))
        {
            ids[source.Op] = nextGroup;
            ids[destination.Op] = nextGroup++;
        }
        return script with { Operations = Rewrite(script.Operations, ids) };
    }

    /// <summary>Whitespace runs collapsed to one space, ends trimmed; and the Word-token count. A token with no
    /// text of its own (image, note reference, opaque content) contributes its match key, so two paragraphs
    /// that differ only in which image they hold never read as the same text.</summary>
    internal static (string Text, int Words) Normalize(IReadOnlyList<IrDiffToken> tokens) =>
        Normalize(tokens, 0, tokens.Count);

    /// <inheritdoc cref="Normalize(IReadOnlyList{IrDiffToken})"/>
    internal static (string Text, int Words) Normalize(IReadOnlyList<IrDiffToken> tokens, int start, int end)
    {
        var sb = new StringBuilder();
        int words = 0;
        bool space = false;
        for (int i = start; i < end; i++)
        {
            var token = tokens[i];
            if (token.Kind == IrDiffTokenKind.Word)
                words++;
            if (token.Kind is IrDiffTokenKind.Image or IrDiffTokenKind.NoteRef or IrDiffTokenKind.Opaque or
                IrDiffTokenKind.Textbox)
            {
                if (space)
                    sb.Append(' ');
                space = false;
                sb.Append('\u0001').Append(token.MatchKey).Append('\u0001');
                continue;
            }
            foreach (char c in token.Text)
            {
                if (char.IsWhiteSpace(c))
                {
                    space = sb.Length > 0;
                    continue;
                }
                if (space)
                    sb.Append(' ');
                space = false;
                sb.Append(c);
            }
        }
        return (sb.ToString(), words);
    }

    /// <summary>The block lists of the cells inside one op's table diff that both surfaces draw cell by cell:
    /// an in-place Modified table that does not take the whole-table fallback. A moved table is drawn whole;
    /// rows that are themselves inserted, deleted or moved are whole-row changes with no cell lists.</summary>
    private static IEnumerable<IrNodeList<IrEditOp>> CellOpLists(IrEditOp op, IrDiffSettings settings) =>
        op.Kind == IrEditOpKind.ModifyBlock && op.TableDiff is { } table &&
        !IrTableDiffer.NeedsWholeTableFallback(table, settings)
            ? table.RowOps
                .SelectMany(row => row.CellOps ?? Enumerable.Empty<IrCellOp>())
                .Where(cell => cell.BlockOps is not null)
                .Select(cell => cell.BlockOps!)
            : Enumerable.Empty<IrNodeList<IrEditOp>>();

    /// <summary>The largest move group id anywhere in the script. Ids are unique across scopes (issue #924),
    /// so relocation ids continue above it.</summary>
    private static int MaxMoveGroupId(IrEditScript script)
    {
        int max = 0;
        void Visit(IEnumerable<IrEditOp> ops)
        {
            foreach (var op in ops)
            {
                max = System.Math.Max(max, op.MoveGroupId ?? 0);
                foreach (var row in op.TableDiff?.RowOps ?? Enumerable.Empty<IrRowOp>())
                {
                    max = System.Math.Max(max, row.MoveGroupId ?? 0);
                    foreach (var cell in row.CellOps ?? Enumerable.Empty<IrCellOp>())
                        Visit(cell.BlockOps ?? Enumerable.Empty<IrEditOp>());
                }
                foreach (var box in op.TextboxDiffs ?? Enumerable.Empty<IrTextboxDiff>())
                    Visit(box.Ops);
            }
        }
        Visit(script.Operations);
        foreach (var note in script.NoteOps ?? Enumerable.Empty<IrNoteDiff>())
            Visit(note.Ops);
        foreach (var story in script.HeaderFooterOps ?? Enumerable.Empty<IrHeaderFooterDiff>())
            Visit(story.Ops);
        return max;
    }

    private static IrNodeList<IrEditOp> Rewrite(IrNodeList<IrEditOp> ops, Dictionary<IrEditOp, int> ids) =>
        IrNodeList.From(ops.Select(op =>
        {
            var rewritten = ids.TryGetValue(op, out var id) ? op with { RelocationGroupId = id } : op;
            if (op.TableDiff is not { } table)
                return rewritten;
            return rewritten with
            {
                TableDiff = new IrTableDiff(IrNodeList.From(table.RowOps.Select(row => row.CellOps is { } cells
                    ? row with
                    {
                        CellOps = IrNodeList.From(cells.Select(c =>
                            c.BlockOps is { } blockOps ? c with { BlockOps = Rewrite(blockOps, ids) } : c)),
                    }
                    : row))),
            };
        }));
}
