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

    /// <summary>The paragraph's WORD-token MatchKey multiset (key → multiplicity). Cached per Align
    /// call; used by the junction pass's uniqueness discipline.</summary>
    public IReadOnlyDictionary<string, int> WordKeys(IrParagraph paragraph) => Bag(paragraph).WordCounts;

    /// <summary>
    /// The subset of a paragraph's word-token keys that contains lexical content (at least one
    /// letter).  The weak, corpus-calibrated junction matcher uses this rather than every
    /// word-token: a shared ordinal such as <c>17</c> is positional scaffolding, not evidence that
    /// two otherwise unrelated paragraphs are revisions of one another.  Alphanumeric labels such
    /// as <c>Q1</c> remain lexical and therefore continue to count.
    /// </summary>
    public IReadOnlyDictionary<string, int> PairingWordKeys(IrParagraph paragraph) => Bag(paragraph).PairingWordCounts;

    /// <summary>Number of lexical word tokens used by the weak junction-pairing evidence gate.</summary>
    public int PairingWordCount(IrParagraph paragraph) => Bag(paragraph).PairingWordCount;

    /// <summary>
    /// The weak junction matcher admits lexical words plus a deliberately narrow semantic-numeric
    /// category: four-digit calendar years. A year carries enough meaning to preserve Word's
    /// date↔heading pairing, unlike a list ordinal such as <c>17</c>; the stricter
    /// <see cref="PairingWordKeys"/> gate remains unchanged for similarity pairing.
    /// </summary>
    public IReadOnlyDictionary<string, int> JunctionWordKeys(IrParagraph paragraph) => Bag(paragraph).JunctionWordCounts;

    /// <summary>Number of lexical-or-year tokens used by the weak junction matcher.</summary>
    public int JunctionWordCount(IrParagraph paragraph) => Bag(paragraph).JunctionWordCount;

    /// <summary>
    /// Should a lone 1×1 gap residue of these two paragraphs force-pair as Modified? True when
    /// EITHER side has no <see cref="IrDiffTokenKind.Word"/> tokens at all (an atomic-only or empty
    /// paragraph — textboxes/images carry no lexical evidence to demand, and demoting them to
    /// Delete+Insert loses the nested textbox/image diff), or when the two sides share at least one
    /// word by RAW TEXT — punctuation-trimmed and normalized according to
    /// <see cref="IrDiffSettings.CaseInsensitive"/>, so "This." shares "This", and a
    /// hyperlink word whose target changed (different MatchKey link suffix) still counts as the
    /// same word. This is deliberately laxer than <see cref="WordOverlap"/>'s MatchKey grain: the
    /// residue test asks "is this the same block, edited?", not "do these tokens diff Equal?".
    /// A full rewrite (zero shared trimmed words, both sides lexical) returns false — the Word
    /// oracle keeps those as separate ins/del paragraphs ("24" ↔ "1.5 Line Spacing Demo").
    /// </summary>
    public bool ResidueForcePair(IrParagraph left, IrParagraph right)
    {
        var a = Bag(left);
        var b = Bag(right);
        if (a.WordCount == 0 || b.WordCount == 0)
            return true;
        // A single word replaced by a single word is "the same short label, edited" — a typo
        // ("Nested." → "Nexted.", WC043's cell), a renumbering ("Two" → "Two1", RC-0010), a
        // retargeted link text — WmlComparer pairs these positionally, and demoting them loses the
        // compat token grain (and turned RC-0010's disjoint-reviewer composition into a false
        // conflict). The oracle's kept-separate cases are all one-word-vs-MULTI-word ("24" ↔
        // "1.5 Line Spacing Demo"), which this rule does not touch.
        if (a.WordCount == 1 && b.WordCount == 1)
            return true;
        var (small, large) = a.TrimmedWords.Count <= b.TrimmedWords.Count ? (a, b) : (b, a);
        foreach (var w in small.TrimmedWords)
            if (large.TrimmedWords.Contains(w))
                return true;
        return false;
    }

    /// <summary>
    /// WORD-token overlap statistics for a paragraph pair: the multiset intersection size over
    /// <see cref="IrDiffTokenKind.Word"/>-kind MatchKeys only, and the Jaccard index over those
    /// word-only multisets. Separator/punctuation/atomic tokens are EXCLUDED — a shared "." or a
    /// run of shared whitespace is no evidence two paragraphs correspond (decoded from the
    /// Word-compare oracle corpus: Word never pairs paragraphs on punctuation-only overlap).
    /// Empty or whitespace-only paragraphs have zero shared words, so they can never qualify.
    /// Uses the same per-Align-call bag cache as <see cref="Score"/>.
    /// </summary>
    public (int SharedWords, double WordJaccard) WordOverlap(IrParagraph left, IrParagraph right)
    {
        var a = Bag(left);
        var b = Bag(right);
        if (a.WordCount == 0 || b.WordCount == 0)
            return (0, 0.0);

        int intersection = 0;
        var (small, large) = a.WordCounts.Count <= b.WordCounts.Count ? (a, b) : (b, a);
        foreach (var kv in small.WordCounts)
            if (large.WordCounts.TryGetValue(kv.Key, out int other))
                intersection += System.Math.Min(kv.Value, other);

        int union = a.WordCount + b.WordCount - intersection;
        return (intersection, union == 0 ? 0.0 : (double)intersection / union);
    }

    /// <summary>
    /// Lexical-only variant of <see cref="WordOverlap"/> for the weak junction-pairing pass.
    /// Numeric-only tokens are deliberately excluded: otherwise synthetic or numbered documents
    /// whose paragraphs merely share their ordinal become an all-Modified diagonal despite having
    /// no common prose.
    /// </summary>
    public (int SharedWords, double WordJaccard) PairingWordOverlap(IrParagraph left, IrParagraph right)
    {
        var a = Bag(left);
        var b = Bag(right);
        if (a.PairingWordCount == 0 || b.PairingWordCount == 0)
            return (0, 0.0);

        int intersection = 0;
        var (small, large) = a.PairingWordCounts.Count <= b.PairingWordCounts.Count ? (a, b) : (b, a);
        foreach (var kv in small.PairingWordCounts)
            if (large.PairingWordCounts.TryGetValue(kv.Key, out int other))
                intersection += System.Math.Min(kv.Value, other);

        int union = a.PairingWordCount + b.PairingWordCount - intersection;
        return (intersection, union == 0 ? 0.0 : (double)intersection / union);
    }

    /// <summary>
    /// Junction-only overlap over lexical word tokens plus four-digit calendar years. This preserves
    /// the numeric-ordinal protection for the general similarity pass while allowing Word's known
    /// date-bearing weak pairing behavior to participate in the junction LCS.
    /// </summary>
    public (int SharedWords, double WordJaccard) JunctionWordOverlap(IrParagraph left, IrParagraph right)
    {
        var a = Bag(left);
        var b = Bag(right);
        if (a.JunctionWordCount == 0 || b.JunctionWordCount == 0)
            return (0, 0.0);

        int intersection = 0;
        var (small, large) = a.JunctionWordCounts.Count <= b.JunctionWordCounts.Count ? (a, b) : (b, a);
        foreach (var kv in small.JunctionWordCounts)
            if (large.JunctionWordCounts.TryGetValue(kv.Key, out int other))
                intersection += System.Math.Min(kv.Value, other);

        int union = a.JunctionWordCount + b.JunctionWordCount - intersection;
        return (intersection, union == 0 ? 0.0 : (double)intersection / union);
    }

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

    /// <summary>A token-MatchKey multiset plus the Word-kind token count (and word-only sub-multiset),
    /// built once per block.</summary>
    private sealed class MatchKeyBag
    {
        private static readonly Dictionary<string, int> EmptyCounts = new();

        private static readonly HashSet<string> EmptyWords = new();

        public Dictionary<string, int> Counts { get; }
        public Dictionary<string, int> WordCounts { get; }  // Word-kind tokens only, by MatchKey
        public Dictionary<string, int> PairingWordCounts { get; } // Word keys containing at least one letter
        public Dictionary<string, int> JunctionWordCounts { get; } // lexical keys plus year-like values
        public HashSet<string> TrimmedWords { get; }        // raw word texts, punct-trimmed + case-folded
        public int Total { get; }      // sum of all multiplicities (every token kind)
        public int WordCount { get; }  // Word-kind tokens only
        public int PairingWordCount { get; } // Word-kind tokens carrying lexical (not ordinal-only) content
        public int JunctionWordCount { get; } // Lexical tokens plus semantic calendar years

        private MatchKeyBag(Dictionary<string, int> counts, Dictionary<string, int> wordCounts,
            Dictionary<string, int> pairingWordCounts, Dictionary<string, int> junctionWordCounts,
            HashSet<string> trimmedWords, int total, int wordCount, int pairingWordCount,
            int junctionWordCount)
        {
            Counts = counts;
            WordCounts = wordCounts;
            PairingWordCounts = pairingWordCounts;
            JunctionWordCounts = junctionWordCounts;
            TrimmedWords = trimmedWords;
            Total = total;
            WordCount = wordCount;
            PairingWordCount = pairingWordCount;
            JunctionWordCount = junctionWordCount;
        }

        public static MatchKeyBag Build(IrParagraph paragraph, IrDiffSettings settings)
        {
            var tokens = IrDiffTokenizer.Tokenize(paragraph, settings);
            var counts = new Dictionary<string, int>();
            var wordCounts = new Dictionary<string, int>();
            var pairingWordCounts = new Dictionary<string, int>();
            var junctionWordCounts = new Dictionary<string, int>();
            var trimmedWords = new HashSet<string>();
            int wordCount = 0, pairingWordCount = 0, junctionWordCount = 0;
            foreach (var t in tokens)
            {
                counts[t.MatchKey] = counts.TryGetValue(t.MatchKey, out int c) ? c + 1 : 1;
                if (t.Kind == IrDiffTokenKind.Word)
                {
                    wordCounts[t.MatchKey] = wordCounts.TryGetValue(t.MatchKey, out int w) ? w + 1 : 1;
                    wordCount++;
                    if (ContainsLetter(t.Text))
                    {
                        pairingWordCounts[t.MatchKey] = pairingWordCounts.TryGetValue(t.MatchKey, out int p) ? p + 1 : 1;
                        pairingWordCount++;
                    }
                    if (ContainsLetter(t.Text) || IsCalendarYear(t.Text))
                    {
                        junctionWordCounts[t.MatchKey] = junctionWordCounts.TryGetValue(t.MatchKey, out int j) ? j + 1 : 1;
                        junctionWordCount++;
                    }
                    AddLexicalPieces(t.Text, trimmedWords, settings);
                }
            }
            return new MatchKeyBag(
                counts, wordCounts, pairingWordCounts, junctionWordCounts, trimmedWords, tokens.Count,
                wordCount, pairingWordCount, junctionWordCount);
        }

        private static bool ContainsLetter(string value)
        {
            foreach (char c in value)
                if (char.IsLetter(c))
                    return true;
            return false;
        }

        private static bool IsCalendarYear(string value)
        {
            if (value.Length != 4)
                return false;
            int year = 0;
            foreach (char c in value)
            {
                if (c < '0' || c > '9')
                    return false;
                year = year * 10 + (c - '0');
            }
            return year is >= 1900 and <= 2099;
        }

        /// <summary>Split a raw word on EVERY non-letter/digit character and add pieces normalized
        /// under the configured case policy — the lexical identity the 1×1-residue evidence test compares on.
        /// Word-style boundaries: "This." contributes "this"; "www.ericwhite.com" contributes
        /// "www"/"ericwhite"/"com", so a hyperlink whose target text changed one segment still
        /// shares lexical content with its original.</summary>
        private static void AddLexicalPieces(string raw, HashSet<string> sink, IrDiffSettings settings)
        {
            int start = -1;
            for (int i = 0; i <= raw.Length; i++)
            {
                bool wordChar = i < raw.Length && char.IsLetterOrDigit(raw[i]);
                if (wordChar)
                {
                    if (start < 0)
                        start = i;
                }
                else if (start >= 0)
                {
                    string piece = raw.Substring(start, i - start);
                    if (settings.CaseInsensitive)
                        piece = settings.Culture is { } culture
                            ? piece.ToLower(culture)
                            : piece.ToLowerInvariant();
                    sink.Add(piece);
                    start = -1;
                }
            }
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
            // Table bags never feed WordOverlap/ResidueForcePair (junction pairing is
            // paragraph-only), so the word-only structures are not materialized.
            return new MatchKeyBag(
                counts, EmptyCounts, EmptyCounts, EmptyCounts, EmptyWords, total, wordCount, 0, 0);
        }
    }
}
