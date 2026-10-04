// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Linq;
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
/// <para>A pair qualifies under the aligner's own fuzzy-move rule: both paragraphs carry at least
/// <see cref="IrDiffSettings.MoveMinimumTokenCount"/> words and score at least
/// <see cref="IrDiffSettings.MoveSimilarityThreshold"/>. Like a moved-and-edited paragraph, the halves are
/// drawn whole, so a lightly edited relocation is still exact under accept and reject.</para>
/// <para>Only the two-way comparison calls this; Consolidate's per-reviewer scripts never carry relocations.</para>
/// </summary>
internal static class IrRelocationPairer
{
    /// <summary>Above this many candidate (deleted, inserted) pairs, only exact-content pairs are scored,
    /// keeping the pass linear on long documents with many unrelated changes.</summary>
    private const long FuzzyPairBudget = 250_000;

    private sealed record Candidate(IrEditOp Op, int ListId, int Order, IrParagraph Paragraph);

    public static IrEditScript Apply(IrEditScript script, IrDocument left, IrDocument right, IrDiffSettings settings)
    {
        var deleted = new List<Candidate>();
        var inserted = new List<Candidate>();
        int nextList = 0;
        int order = 0;
        void Collect(IEnumerable<IrEditOp> ops)
        {
            int listId = nextList++;
            foreach (var op in ops)
            {
                if (op.Kind == IrEditOpKind.DeleteBlock && Resolve(left, op.LeftAnchor) is IrParagraph lp)
                    deleted.Add(new Candidate(op, listId, order++, lp));
                else if (op.Kind == IrEditOpKind.InsertBlock && Resolve(right, op.RightAnchor) is IrParagraph rp)
                    inserted.Add(new Candidate(op, listId, order++, rp));
                foreach (var cellOps in CellOpLists(op))
                    Collect(cellOps);
            }
        }
        Collect(script.Operations);
        if (deleted.Count == 0 || inserted.Count == 0 || nextList == 1)
            return script;

        var groups = Pair(deleted, inserted, settings);
        if (groups.Count == 0)
            return script;

        int nextGroup = MaxMoveGroupId(script) + 1;
        var ids = new Dictionary<IrEditOp, int>(ReferenceEqualityComparer.Instance);
        foreach (var (source, destination) in groups.OrderBy(g => g.Destination.Order))
        {
            ids[source.Op] = nextGroup;
            ids[destination.Op] = nextGroup++;
        }
        return script with { Operations = Rewrite(script.Operations, ids) };
    }

    /// <summary>Greedy best-first pairing across DIFFERENT block lists, ties broken by document order.</summary>
    private static List<(Candidate Source, Candidate Destination)> Pair(
        List<Candidate> deleted, List<Candidate> inserted, IrDiffSettings settings)
    {
        var similarity = new IrBlockSimilarity(settings);
        bool fuzzy = (long)deleted.Count * inserted.Count <= FuzzyPairBudget;
        var scored = new List<(double Score, Candidate Source, Candidate Destination)>();
        if (fuzzy)
        {
            foreach (var d in deleted)
            {
                if (similarity.WordCount(d.Paragraph) < settings.MoveMinimumTokenCount)
                    continue;
                foreach (var i in inserted)
                {
                    if (i.ListId == d.ListId || similarity.WordCount(i.Paragraph) < settings.MoveMinimumTokenCount)
                        continue;
                    double score = similarity.Score(d.Paragraph, i.Paragraph);
                    if (score >= settings.MoveSimilarityThreshold)
                        scored.Add((score, d, i));
                }
            }
        }
        else
        {
            var byHash = inserted.ToLookup(i => i.Paragraph.ContentHash);
            foreach (var d in deleted)
            {
                if (similarity.WordCount(d.Paragraph) < settings.MoveMinimumTokenCount)
                    continue;
                foreach (var i in byHash[d.Paragraph.ContentHash])
                    if (i.ListId != d.ListId)
                        scored.Add((1.0, d, i));
            }
        }

        var used = new HashSet<Candidate>(ReferenceEqualityComparer.Instance);
        var pairs = new List<(Candidate, Candidate)>();
        foreach (var (_, source, destination) in scored
                     .OrderByDescending(s => s.Score)
                     .ThenBy(s => s.Source.Order)
                     .ThenBy(s => s.Destination.Order))
        {
            if (used.Contains(source) || used.Contains(destination))
                continue;
            used.Add(source);
            used.Add(destination);
            pairs.Add((source, destination));
        }
        return pairs;
    }

    private static IrBlock? Resolve(IrDocument doc, string? anchor) =>
        anchor is not null && doc.AnchorIndex.TryGetValue(anchor, out var block) ? block : null;

    /// <summary>The block lists of the paired cells inside one op's table diff (rows that are themselves
    /// inserted, deleted or moved are whole-row changes with no cell lists).</summary>
    private static IEnumerable<IrNodeList<IrEditOp>> CellOpLists(IrEditOp op) =>
        (op.TableDiff?.RowOps ?? Enumerable.Empty<IrRowOp>())
            .SelectMany(row => row.CellOps ?? Enumerable.Empty<IrCellOp>())
            .Where(cell => cell.BlockOps is not null)
            .Select(cell => cell.BlockOps!);

    /// <summary>The largest move group id anywhere in the script. Ids are unique across scopes, so
    /// relocation ids continue above it.</summary>
    private static int MaxMoveGroupId(IrEditScript script)
    {
        int max = 0;
        void Visit(IEnumerable<IrEditOp> ops)
        {
            foreach (var op in ops)
            {
                max = System.Math.Max(max, op.MoveGroupId ?? 0);
                foreach (var row in op.TableDiff?.RowOps ?? Enumerable.Empty<IrRowOp>())
                    max = System.Math.Max(max, row.MoveGroupId ?? 0);
                foreach (var cellOps in CellOpLists(op))
                    Visit(cellOps);
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
