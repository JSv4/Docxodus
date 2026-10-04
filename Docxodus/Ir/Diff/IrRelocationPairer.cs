// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Linq;
using System.Text;
using Docxodus.Ir;

namespace Docxodus.Ir.Diff;

/// <summary>
/// Pairs relocated content the aligner cannot see as a move, because it detects moves within one block list
/// (the body, or one table cell) and only between whole blocks. This pass gives the two halves of each such
/// relocation a shared relocation id and changes nothing else: every op and token span keeps its kind, so
/// applying the script, and accepting or rejecting its markup, is exactly what it was. The renderers draw a
/// tagged pair as <c>w:moveFrom</c>/<c>w:moveTo</c>, and <c>GetRevisions</c> reports it as a Moved pair, when
/// move reporting is on. Two kinds of relocation are paired:
/// <list type="bullet">
/// <item><b>Whole paragraphs across a table boundary</b> (issue #887): a <see cref="IrEditOpKind.DeleteBlock"/>
/// paragraph in one block list and an <see cref="IrEditOpKind.InsertBlock"/> paragraph in another (the body and
/// a cell, or two cells). Tagged with <see cref="IrEditOp.RelocationGroupId"/>.</item>
/// <item><b>Text spans</b> (issue #888): a deleted span inside a modified paragraph whose text reappears as an
/// inserted span elsewhere — later in the same paragraph, in another paragraph, or in a table cell — or as a
/// whole inserted paragraph, and the reverse (a whole deleted paragraph whose text now sits inside another).
/// Tagged with <see cref="IrTokenOp.RelocationGroupId"/> on a span, and <see cref="IrEditOp.RelocationGroupId"/>
/// on a whole-paragraph half.</item>
/// </list>
/// <para><b>What pairs.</b> The two texts must be identical once whitespace runs are collapsed, word for word in
/// the same run formatting, and carry at least <see cref="IrDiffSettings.MoveMinimumTokenCount"/> words. An
/// exact match leaves no in-move edit — text or formatting — for either surface to describe. A deleted and an inserted span in the same paragraph pair only when retained
/// words separate them; adjacent, they are a replacement. Each half pairs at most once.</para>
/// <para><b>Boundary slides.</b> A token diff may place a span's edge differently from where the moved text
/// begins: deleting "first sentence. The " after a retained "The " deletes the same text as "The first
/// sentence. " before a retained "The ". Before comparing texts, each span is also considered at every
/// position it can slide to without changing what it deletes or inserts, and the chosen position is written
/// back into the token diff.</para>
/// <para><b>Where it looks.</b> The body and the cells of tables drawn cell by cell on both surfaces
/// (<see cref="IrTableDiffer.NeedsWholeTableFallback"/> false); never a moved table. A paragraph the redline
/// draws inside a cross-paragraph run (<see cref="IrEditOpKind.CrossParagraphRunBlock"/>) is left out — the
/// markup script has no per-paragraph diff for it, and the data script is told to leave out the same
/// paragraphs — so the two scripts pair alike. A paragraph the markup renderer always draws whole (an
/// inseparable carrier, a textbox interior) offers no spans. With
/// <see cref="IrDiffSettings.PreserveInputRevisions"/> on, nothing pairs: the markup renderer may draw a block
/// whole to keep its input revisions, which the revision list cannot see. Nor with move reporting off, or under
/// the WmlComparer-compatible revision grain: the script is then left exactly as the builder made it.</para>
/// <para>Only the two-way comparison calls this; Consolidate's per-reviewer scripts never carry relocations.</para>
/// </summary>
internal static class IrRelocationPairer
{
    /// <summary>How far a span edge is allowed to slide in each direction.</summary>
    private const int MaxSlide = 32;

    /// <summary>A whole deleted or inserted paragraph.</summary>
    private sealed record Candidate(IrEditOp Op, int ListId, string Text, IReadOnlyList<IrRunFormat?> Formats, bool IsSource);

    /// <summary>One token-level step of a diff: an aligned pair (Equal/FormatChanged), a deleted left token
    /// or an inserted right token. Index -1 = the side the step does not consume.</summary>
    private readonly record struct Atom(IrTokenOpKind Kind, int L, int R, int? Relocation);

    /// <summary>A modified paragraph's token diff, which can carry relocated spans.</summary>
    private sealed class Host
    {
        public IrEditOp Op { get; init; } = null!;
        public int ListId { get; init; }
        public IReadOnlyList<IrDiffToken> Left { get; init; } = System.Array.Empty<IrDiffToken>();
        public IReadOnlyList<IrDiffToken> Right { get; init; } = System.Array.Empty<IrDiffToken>();
        public List<Atom> Atoms { get; init; } = new();
        public IrFormatComparison FormatComparison { get; init; }
        public HashSet<int> Claimed { get; } = new();
        public bool Changed { get; set; }
    }

    /// <summary>A span half: a Delete/Insert run of atoms [AtomStart, AtomEnd) in a host, or a whole
    /// paragraph (<see cref="Block"/>).</summary>
    private sealed class Span
    {
        public Host? Host { get; init; }
        public int AtomStart { get; init; }
        public int AtomEnd { get; init; }
        public Candidate? Block { get; init; }
        public bool IsSource { get; init; }
        public int Order { get; init; }
        public List<(int Offset, string Text, IReadOnlyList<IrRunFormat?> Formats)> Variants { get; } = new();
        public bool Used { get; set; }
    }

    /// <param name="fusedAnchors">Body paragraph anchors the redline draws inside cross-paragraph runs, when
    /// this script is the data script of a comparison whose redline fuses runs; null otherwise.</param>
    public static IrEditScript Apply(
        IrEditScript script, IrDocument left, IrDocument right, IrDiffSettings settings,
        IReadOnlySet<string>? fusedAnchors = null)
    {
        // Pairing only labels content as moved, so it runs only where a move is drawn and reported: move
        // reporting on, the engine's fine revision grain (the WmlComparer-compatible grain reproduces that
        // comparer's revision set, which has no such moves), and no input revisions being preserved.
        if (settings.PreserveInputRevisions || !settings.RenderMoves ||
            settings.RevisionGranularity != RevisionGranularity.Fine)
            return script;

        int min = settings.MoveMinimumTokenCount;
        var blocks = new List<Candidate>();
        var hosts = new List<Host>();
        int nextList = 0;
        bool Fused(string? anchor) => anchor is not null && fusedAnchors?.Contains(anchor) == true;
        void Collect(IEnumerable<IrEditOp> ops)
        {
            int listId = nextList++;
            foreach (var op in ops)
            {
                if (op.Kind == IrEditOpKind.DeleteBlock && Paragraph(op.LeftAnchor, left) is { } lp)
                    AddBlock(op, lp, listId, isSource: true);
                else if (op.Kind == IrEditOpKind.InsertBlock && Paragraph(op.RightAnchor, right) is { } rp)
                    AddBlock(op, rp, listId, isSource: false);
                else if (op.Kind == IrEditOpKind.ModifyBlock && op.TokenDiff is { } diff &&
                         !op.RequiresWholeParagraphReplace && op.TextboxDiffs is null &&
                         Paragraph(op.LeftAnchor, left) is { } ml && Paragraph(op.RightAnchor, right) is { } mr)
                    hosts.Add(new Host
                    {
                        Op = op, ListId = listId, Atoms = Atomize(diff), FormatComparison = settings.FormatComparison,
                        Left = IrDiffTokenizer.Tokenize(ml, settings), Right = IrDiffTokenizer.Tokenize(mr, settings),
                    });
                foreach (var cellOps in CellOpLists(op, settings))
                    Collect(cellOps);
            }
        }
        IrParagraph? Paragraph(string? anchor, IrDocument doc) =>
            anchor is not null && !Fused(anchor) && doc.AnchorIndex.TryGetValue(anchor, out var block) &&
            block is IrParagraph paragraph ? paragraph : null;
        void AddBlock(IrEditOp op, IrParagraph paragraph, int listId, bool isSource)
        {
            var tokens = IrDiffTokenizer.Tokenize(paragraph, settings);
            var (text, words) = Normalize(tokens);
            if (words >= min)
                blocks.Add(new Candidate(op, listId, text, WordFormats(tokens, 0, tokens.Count), isSource));
        }
        Collect(script.Operations);
        if (blocks.Count == 0 && hosts.Count == 0)
            return script;

        var groupOf = new Dictionary<IrEditOp, int>(ReferenceEqualityComparer.Instance);
        int nextGroup = MaxMoveGroupId(script) + 1;

        // 1. Whole paragraphs across a table boundary (issue #887).
        var paired = new HashSet<Candidate>(ReferenceEqualityComparer.Instance);
        var byText = blocks.Where(b => !b.IsSource).GroupBy(b => b.Text, System.StringComparer.Ordinal)
            .ToDictionary(g => g.Key, g => g.ToList(), System.StringComparer.Ordinal);
        var blockPairs = new List<(Candidate Source, Candidate Destination)>();
        foreach (var source in blocks.Where(b => b.IsSource))
        {
            if (!byText.TryGetValue(source.Text, out var destinations))
                continue;
            var destination = destinations.FirstOrDefault(d => d.ListId != source.ListId && !paired.Contains(d) &&
                SameFormats(source.Formats, d.Formats, settings.FormatComparison));
            if (destination is null)
                continue;
            paired.Add(source);
            paired.Add(destination);
            blockPairs.Add((source, destination));
        }
        var blockOrder = new Dictionary<Candidate, int>(ReferenceEqualityComparer.Instance);
        for (int i = 0; i < blocks.Count; i++)
            blockOrder[blocks[i]] = i;
        foreach (var (source, destination) in blockPairs.OrderBy(p => blockOrder[p.Destination]))
        {
            groupOf[source.Op] = nextGroup;
            groupOf[destination.Op] = nextGroup++;
        }

        // 2. Text spans, against each other and against the whole paragraphs still unpaired (issue #888).
        PairSpans(hosts, blocks.Where(b => !paired.Contains(b)).ToList(), groupOf, nextGroup, min, settings.FormatComparison);

        var replaced = new Dictionary<IrEditOp, IrEditOp>(ReferenceEqualityComparer.Instance);
        foreach (var host in hosts.Where(h => h.Changed))
            replaced[host.Op] = host.Op with { TokenDiff = new IrTokenDiff(Regroup(host.Atoms)) };
        foreach (var (op, id) in groupOf)
            replaced[op] = (replaced.TryGetValue(op, out var r) ? r : op) with { RelocationGroupId = id };
        return replaced.Count == 0 ? script : script with { Operations = Rewrite(script.Operations, replaced) };
    }

    // ------------------------------------------------------------------ text spans (issue #888)

    private static void PairSpans(
        List<Host> hosts, List<Candidate> blocks, Dictionary<IrEditOp, int> groupOf, int nextGroup, int min,
        IrFormatComparison formatComparison)
    {
        var spans = new List<Span>();
        int order = 0;
        foreach (var host in hosts)
        {
            int a = 0;
            while (a < host.Atoms.Count)
            {
                var kind = host.Atoms[a].Kind;
                int b = a + 1;
                while (b < host.Atoms.Count && host.Atoms[b].Kind == kind)
                    b++;
                if (kind is IrTokenOpKind.Delete or IrTokenOpKind.Insert)
                {
                    var span = new Span
                    {
                        Host = host, AtomStart = a, AtomEnd = b, IsSource = kind == IrTokenOpKind.Delete, Order = order++,
                    };
                    AddVariants(span, min);
                    if (span.Variants.Count > 0)
                        spans.Add(span);
                }
                a = b;
            }
        }
        foreach (var block in blocks)
        {
            var span = new Span { Block = block, IsSource = block.IsSource, Order = order++ };
            span.Variants.Add((0, block.Text, block.Formats));
            spans.Add(span);
        }

        // Destinations by text; each source, in document order, claims an unused destination with the same
        // text in the same formatting. Preferred, in order: edges on a sentence end (so a move reads "The second
        // sentence." rather than "second sentence. The" when both halves can slide), the smallest total slide,
        // the earliest destination.
        var destinations = new Dictionary<string, List<(Span Span, int Offset, IReadOnlyList<IrRunFormat?> Formats)>>(
            System.StringComparer.Ordinal);
        foreach (var span in spans.Where(s => !s.IsSource))
            foreach (var (offset, text, formats) in span.Variants)
            {
                if (!destinations.TryGetValue(text, out var list))
                    destinations[text] = list = new List<(Span, int, IReadOnlyList<IrRunFormat?>)>();
                list.Add((span, offset, formats));
            }

        foreach (var source in spans.Where(s => s.IsSource).OrderBy(s => s.Order))
        {
            (Span Span, int SourceOffset, int DestinationOffset, (int Edge, int Slide, int Order) Key)? best = null;
            foreach (var (offset, text, formats) in source.Variants)
            {
                if (!destinations.TryGetValue(text, out var candidates))
                    continue;
                int edge = EndsSentence(text) ? 0 : 1;
                foreach (var (destination, destinationOffset, destinationFormats) in candidates)
                {
                    if (destination.Used || (source.Block != null && destination.Block != null) ||
                        !SameFormats(formats, destinationFormats, formatComparison) ||
                        !Separated(source, offset, destination, destinationOffset) ||
                        !CanClaim(source, offset) || !CanClaim(destination, destinationOffset))
                        continue;
                    var key = (edge, System.Math.Abs(offset) + System.Math.Abs(destinationOffset), destination.Order);
                    if (best is null || key.CompareTo(best.Value.Key) < 0)
                        best = (destination, offset, destinationOffset, key);
                }
            }
            if (best is not { } chosen)
                continue;
            int id = nextGroup++;
            Commit(source, chosen.SourceOffset, id, groupOf);
            Commit(chosen.Span, chosen.DestinationOffset, id, groupOf);
        }
    }

    /// <summary>A deleted and an inserted span in the SAME token diff are a move only when retained words
    /// separate them; adjacent, they are a replacement.</summary>
    private static bool Separated(Span source, int sourceOffset, Span destination, int destinationOffset)
    {
        if (source.Host is null || !ReferenceEquals(source.Host, destination.Host))
            return true;
        var host = source.Host;
        var (s0, s1) = Window(source, sourceOffset);
        var (d0, d1) = Window(destination, destinationOffset);
        for (int i = System.Math.Min(s1, d1); i < System.Math.Max(s0, d0); i++)
        {
            var atom = host.Atoms[i];
            if (IsAligned(atom) && host.Left[atom.L].Kind == IrDiffTokenKind.Word)
                return true;
        }
        return false;
    }

    /// <summary>The atom positions a span covers, slides included.</summary>
    private static (int Start, int End) Window(Span span, int offset) =>
        offset < 0 ? (span.AtomStart + offset, span.AtomEnd) : (span.AtomStart, span.AtomEnd + offset);

    private static bool CanClaim(Span span, int offset)
    {
        if (span.Host is null)
            return true;
        var (start, end) = Window(span, offset);
        for (int i = start; i < end; i++)
            if (span.Host.Claimed.Contains(i))
                return false;
        return true;
    }

    private static void Commit(Span span, int offset, int id, Dictionary<IrEditOp, int> groupOf)
    {
        span.Used = true;
        if (span.Host is null)
        {
            groupOf[span.Block!.Op] = id;
            return;
        }
        var host = span.Host;
        var (start, end) = Window(span, offset);
        for (int i = start; i < end; i++)
            host.Claimed.Add(i);
        Slide(host, span.AtomStart, span.AtomEnd, offset, span.IsSource);
        for (int i = span.AtomStart + offset; i < span.AtomEnd + offset; i++)
            host.Atoms[i] = host.Atoms[i] with { Relocation = id };
        host.Changed = true;
    }

    /// <summary>Every position a span can slide to, with its normalized text, when that carries enough words.
    /// A slide by one step is allowed when the retained token on the far side matches, in text and match key,
    /// the token the span gives up on the near side, so the span still deletes (or inserts) the same text.</summary>
    private static void AddVariants(Span span, int min)
    {
        var host = span.Host!;
        bool deleted = span.IsSource;
        var tokens = deleted ? host.Left : host.Right;
        int Index(Atom atom) => deleted ? atom.L : atom.R;
        int s = Index(host.Atoms[span.AtomStart]);
        int e = Index(host.Atoms[span.AtomEnd - 1]) + 1;

        void Add(int offset)
        {
            var (text, words) = Normalize(tokens, s + offset, e + offset);
            if (words >= min)
                span.Variants.Add((offset, text, WordFormats(tokens, s + offset, e + offset)));
        }
        Add(0);
        for (int k = 1; k <= MaxSlide; k++)
        {
            int a = span.AtomStart - k;
            if (a < 0 || !IsAligned(host.Atoms[a]) || !Same(tokens[s - k], tokens[e - k]))
                break;
            Add(-k);
        }
        for (int k = 1; k <= MaxSlide; k++)
        {
            int b = span.AtomEnd + k - 1;
            if (b >= host.Atoms.Count || !IsAligned(host.Atoms[b]) || !Same(tokens[s + k - 1], tokens[e + k - 1]))
                break;
            Add(k);
        }
    }

    /// <summary>Slide a span's atoms by <paramref name="offset"/> steps, re-pairing the retained tokens it
    /// passes over. The span still covers the same text, and each re-paired position is Equal or
    /// FormatChanged by the token differ's own format rule.</summary>
    private static void Slide(Host host, int a, int b, int offset, bool deleted)
    {
        if (offset == 0)
            return;
        var atoms = host.Atoms;
        int length = b - a;
        int s = deleted ? atoms[a].L : atoms[a].R;
        var kind = atoms[a].Kind;
        Atom Moved(int index) => deleted ? new Atom(kind, index, -1, null) : new Atom(kind, -1, index, null);
        Atom Repaired(int moved, Atom retained) =>
            deleted ? Pair(host, moved, retained.R) : Pair(host, retained.L, moved);
        if (offset < 0)
        {
            int k = -offset;
            var retained = atoms.GetRange(a - k, k);   // the aligned pairs the span slides over
            for (int j = 0; j < length; j++)
                atoms[a - k + j] = Moved(s - k + j);
            for (int j = 0; j < k; j++)
                atoms[b - k + j] = Repaired(s + length - k + j, retained[j]);
        }
        else
        {
            int k = offset;
            var retained = atoms.GetRange(b, k);
            for (int j = 0; j < k; j++)
                atoms[a + j] = Repaired(s + j, retained[j]);
            for (int j = 0; j < length; j++)
                atoms[a + k + j] = Moved(s + k + j);
        }
    }

    private static Atom Pair(Host host, int l, int r) => new(
        IrModeledFormat.RunFormatEqual(host.Left[l].Format, host.Right[r].Format, host.FormatComparison)
            ? IrTokenOpKind.Equal
            : IrTokenOpKind.FormatChanged,
        l, r, null);

    private static bool IsAligned(Atom atom) => atom.Kind is IrTokenOpKind.Equal or IrTokenOpKind.FormatChanged;

    private static bool Same(IrDiffToken a, IrDiffToken b) =>
        a.Kind == b.Kind && a.MatchKey == b.MatchKey && string.Equals(a.Text, b.Text, System.StringComparison.Ordinal);

    private static List<Atom> Atomize(IrTokenDiff diff)
    {
        var atoms = new List<Atom>();
        foreach (var op in diff.Ops)
        {
            switch (op.Kind)
            {
                case IrTokenOpKind.Delete:
                    for (int l = op.LeftStart; l < op.LeftEnd; l++)
                        atoms.Add(new Atom(op.Kind, l, -1, op.RelocationGroupId));
                    break;
                case IrTokenOpKind.Insert:
                    for (int r = op.RightStart; r < op.RightEnd; r++)
                        atoms.Add(new Atom(op.Kind, -1, r, op.RelocationGroupId));
                    break;
                default:
                    for (int i = 0; i < op.LeftLength; i++)
                        atoms.Add(new Atom(op.Kind, op.LeftStart + i, op.RightStart + i, null));
                    break;
            }
        }
        return atoms;
    }

    /// <summary>Rebuild token ops from atoms: maximal runs of one kind and one relocation id.</summary>
    private static IrNodeList<IrTokenOp> Regroup(List<Atom> atoms)
    {
        var ops = new List<IrTokenOp>();
        int lc = 0, rc = 0, i = 0;
        while (i < atoms.Count)
        {
            var kind = atoms[i].Kind;
            var relocation = atoms[i].Relocation;
            int j = i;
            while (j < atoms.Count && atoms[j].Kind == kind && atoms[j].Relocation == relocation)
                j++;
            int n = j - i;
            ops.Add(kind switch
            {
                IrTokenOpKind.Delete => new IrTokenOp(kind, lc, lc + n, rc, rc, relocation),
                IrTokenOpKind.Insert => new IrTokenOp(kind, lc, lc, rc, rc + n, relocation),
                _ => new IrTokenOp(kind, lc, lc + n, rc, rc + n),
            });
            if (kind != IrTokenOpKind.Insert)
                lc += n;
            if (kind != IrTokenOpKind.Delete)
                rc += n;
            i = j;
        }
        return IrNodeList.From(ops);
    }

    // ------------------------------------------------------------------ shared plumbing

    private static bool EndsSentence(string text) => text.Length > 0 && text[^1] is '.' or '!' or '?' or ';' or ':';

    /// <summary>The run format of each word token in [start, end).</summary>
    internal static IReadOnlyList<IrRunFormat?> WordFormats(IReadOnlyList<IrDiffToken> tokens, int start, int end)
    {
        var formats = new List<IrRunFormat?>();
        for (int i = start; i < end; i++)
            if (tokens[i].Kind == IrDiffTokenKind.Word)
                formats.Add(tokens[i].Format);
        return formats;
    }

    /// <summary>Whether two halves carry the same run formatting word for word, under the token differ's own
    /// format rule; a relocation whose formatting also changed is not an exact move.</summary>
    internal static bool SameFormats(
        IReadOnlyList<IrRunFormat?> a, IReadOnlyList<IrRunFormat?> b, IrFormatComparison comparison)
    {
        if (a.Count != b.Count)
            return false;
        for (int i = 0; i < a.Count; i++)
            if (!IrModeledFormat.RunFormatEqual(a[i], b[i], comparison))
                return false;
        return true;
    }

    /// <summary>Whether a paired script carries a relocation on a body-level op or span — the only content a
    /// cross-paragraph run can absorb.</summary>
    internal static bool TouchesBody(IrEditScript script) =>
        script.Operations.Any(op => op.RelocationGroupId is not null ||
            op.TokenDiff?.Ops.Any(o => o.RelocationGroupId is not null) == true);

    /// <summary>Whitespace runs collapsed to one space, ends trimmed; and the Word-token count. A token with no
    /// text of its own (image, note reference, opaque content) contributes its match key, so two texts that
    /// differ only in which image they hold never read as the same.</summary>
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

    private static IrNodeList<IrEditOp> Rewrite(IrNodeList<IrEditOp> ops, Dictionary<IrEditOp, IrEditOp> replaced) =>
        IrNodeList.From(ops.Select(op =>
        {
            var rewritten = replaced.TryGetValue(op, out var r) ? r : op;
            if (op.TableDiff is not { } table)
                return rewritten;
            return rewritten with
            {
                TableDiff = new IrTableDiff(IrNodeList.From(table.RowOps.Select(row => row.CellOps is { } cells
                    ? row with
                    {
                        CellOps = IrNodeList.From(cells.Select(c =>
                            c.BlockOps is { } blockOps ? c with { BlockOps = Rewrite(blockOps, replaced) } : c)),
                    }
                    : row))),
            };
        }));
}
