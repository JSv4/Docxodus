#nullable enable

using System.Collections.Generic;
using Docxodus.Ir;

namespace Docxodus.Ir.Diff;

/// <summary>
/// M2.2 Task 3 block-similarity scorer for the aligner's in-gap pairing and cross-gap fuzzy-move
/// detection. Scores a (left, right) block pair in [0, 1] and caches per-block tokenization so the
/// many candidate-pair scorings inside one <see cref="IrBlockAligner.Align"/> call tokenize each block
/// at most once.
/// </summary>
/// <remarks>
/// <para><b>Score model.</b> For two <see cref="IrParagraph"/>s the score is the Jaccard index over
/// their token <c>MatchKey</c> MULTISETS (multiset intersection size / multiset union size), using the
/// SAME <see cref="IrDiffTokenizer"/> the token diff keys on — so similarity is consistent with the
/// downstream token diff (a pair the scorer rates 1.0 token-diffs to all-Equal). Multiset (not set)
/// semantics matter for repeated-word text ("the the the" vs "the the") so duplicate words contribute
/// their multiplicity. An empty-vs-empty paragraph pair scores 1.0 (both token multisets empty — they
/// ARE the same content); empty-vs-nonempty scores 0.</para>
/// <para><b>Non-paragraph blocks.</b> Tables, section breaks and opaque blocks have no token model in
/// this task, so they score 0 UNLESS their <see cref="IrBlock.ContentHash"/> is equal, in which case
/// they score 1.0 (identical content). This deliberately keeps tables OUT of fuzzy pairing here —
/// row/cell-granular table similarity is Task 4 — while still letting an exact-content non-paragraph
/// block participate (e.g. an exact table relocation the anchoring missed).</para>
/// <para><b>Cost.</b> Each <see cref="Score"/> is O(tokens) given cached tokenizations; the cache makes
/// the cost of scoring G² candidate pairs in a gap O(G·tokens) tokenization + O(G²·tokens) scoring,
/// consistent with the aligner's documented G²/2 in-gap bound.</para>
/// </remarks>
internal sealed class IrBlockSimilarity
{
    private readonly IrDiffSettings _settings;

    // Per-Align-call tokenization cache, keyed by block reference identity. A block is tokenized at most
    // once even though it is scored against many candidates. Word-count is cached alongside (cheap, and
    // the cross-gap move gate needs it for every leftover block).
    private readonly Dictionary<IrParagraph, MatchKeyBag> _bagCache =
        new(ReferenceEqualityComparer.Instance);

    // Per-Align-call table-bag cache (M2.4b Workstream C): a table is tokenized to a flattened multiset of
    // ALL its descendant cell-paragraph tokens at most once, even though it is scored against many candidates.
    private readonly Dictionary<IrTable, MatchKeyBag> _tableBagCache =
        new(ReferenceEqualityComparer.Instance);

    public IrBlockSimilarity(IrDiffSettings settings) => _settings = settings;

    /// <summary>
    /// Score the similarity of <paramref name="left"/> and <paramref name="right"/> in [0, 1].
    /// Paragraph pairs: Jaccard over token MatchKey multisets. Non-paragraph (or mixed) pairs: 1.0 iff
    /// ContentHash-equal, else 0.
    /// </summary>
    public double Score(IrBlock left, IrBlock right)
    {
        if (left is IrParagraph lp && right is IrParagraph rp)
            return Jaccard(Bag(lp), Bag(rp));

        // Table-aware similarity (M2.4b Workstream C): score a TABLE pair by the Jaccard index over their
        // CONCATENATED cell-paragraph token multisets — the same token model paragraphs use, flattened over
        // every descendant cell paragraph. This lets the in-gap pairing classify two structurally-similar
        // tables (e.g. the two endnote tables of WC-1750/1760, which differ only in a couple of cell words) as
        // a Modified pair, so IrTableDiffer can produce row/cell-granular edits instead of a whole-table
        // delete+insert. Exact-content tables still score 1.0 (their token multisets are identical), so this
        // never demotes an exact relocation. NB: this is an ALIGNMENT capability addition — it runs in BOTH
        // Fine and compatible modes; the produced table-row/cell ops are the engine's truth either way.
        if (left is IrTable lt && right is IrTable rt)
            return Jaccard(TableBag(lt), TableBag(rt));

        // Other non-paragraph or mixed-kind pairs: only exact content counts.
        return left.ContentHash.Equals(right.ContentHash) ? 1.0 : 0.0;
    }

    /// <summary>Number of <see cref="IrDiffTokenKind.Word"/> tokens in a block (0 for non-paragraphs).</summary>
    public int WordCount(IrBlock block) => block is IrParagraph p ? Bag(p).WordCount : 0;

    private MatchKeyBag Bag(IrParagraph paragraph)
    {
        if (_bagCache.TryGetValue(paragraph, out var bag))
            return bag;
        bag = MatchKeyBag.Build(paragraph, _settings);
        _bagCache[paragraph] = bag;
        return bag;
    }

    private MatchKeyBag TableBag(IrTable table)
    {
        if (_tableBagCache.TryGetValue(table, out var bag))
            return bag;
        bag = MatchKeyBag.BuildTable(table, _settings);
        _tableBagCache[table] = bag;
        return bag;
    }

    /// <summary>
    /// Jaccard index over two token-MatchKey multisets: sum of per-key min counts (intersection) over
    /// sum of per-key max counts (union). Two empty bags score 1.0 (identical empty content).
    /// </summary>
    private static double Jaccard(MatchKeyBag a, MatchKeyBag b)
    {
        if (a.Total == 0 && b.Total == 0)
            return 1.0;

        int intersection = 0;
        // Iterate the smaller bag for the intersection; union derives from totals (|A|+|B|-|A∩B|).
        var (small, large) = a.Counts.Count <= b.Counts.Count ? (a, b) : (b, a);
        foreach (var kv in small.Counts)
            if (large.Counts.TryGetValue(kv.Key, out int other))
                intersection += System.Math.Min(kv.Value, other);

        int union = a.Total + b.Total - intersection;
        return union == 0 ? 1.0 : (double)intersection / union;
    }

    /// <summary>A token-MatchKey multiset plus the Word-kind token count, built once per block.</summary>
    private sealed class MatchKeyBag
    {
        public Dictionary<string, int> Counts { get; }
        public int Total { get; }      // sum of all multiplicities (every token kind)
        public int WordCount { get; }  // Word-kind tokens only

        private MatchKeyBag(Dictionary<string, int> counts, int total, int wordCount)
        {
            Counts = counts;
            Total = total;
            WordCount = wordCount;
        }

        public static MatchKeyBag Build(IrParagraph paragraph, IrDiffSettings settings)
        {
            var tokens = IrDiffTokenizer.Tokenize(paragraph, settings);
            var counts = new Dictionary<string, int>();
            int wordCount = 0;
            foreach (var t in tokens)
            {
                counts[t.MatchKey] = counts.TryGetValue(t.MatchKey, out int c) ? c + 1 : 1;
                if (t.Kind == IrDiffTokenKind.Word)
                    wordCount++;
            }
            return new MatchKeyBag(counts, tokens.Count, wordCount);
        }

        /// <summary>Flatten a table to one MatchKey multiset over EVERY descendant cell paragraph's tokens
        /// (document order), so two structurally-similar tables score by shared cell content.</summary>
        public static MatchKeyBag BuildTable(IrTable table, IrDiffSettings settings)
        {
            var counts = new Dictionary<string, int>();
            int total = 0, wordCount = 0;
            foreach (var row in table.Rows)
                foreach (var cell in row.Cells)
                    foreach (var block in cell.Blocks)
                        if (block is IrParagraph p)
                            foreach (var t in IrDiffTokenizer.Tokenize(p, settings))
                            {
                                counts[t.MatchKey] = counts.TryGetValue(t.MatchKey, out int c) ? c + 1 : 1;
                                total++;
                                if (t.Kind == IrDiffTokenKind.Word)
                                    wordCount++;
                            }
            return new MatchKeyBag(counts, total, wordCount);
        }
    }
}
