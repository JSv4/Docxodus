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
/// draws inside a cross-paragraph run (<see cref="IrEditOpKind.CrossParagraphRunBlock"/>) has no per-paragraph
/// op in the markup script; <see cref="ApplyToBoth"/> pairs a half there only when the run holds it as one
/// deleted (or inserted) stretch of one output paragraph, and tags that stretch in the markup script with the
/// same id, so the two scripts pair alike (issue #930). A paragraph the markup renderer always draws whole (an
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
    private sealed record Candidate(
        IrEditOp Op, int ListId, string Text, IReadOnlyList<IrRunFormat?> Formats, bool IsSource, CellSpan? Cell = null);

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

        /// <summary>Set when the markup script draws this paragraph pair inside a cross-paragraph run: each
        /// span variant must then map onto one stretch of that run's cells (issue #930).</summary>
        public RedlineView? FusedIn { get; init; }
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
        public List<(int Offset, string Text, IReadOnlyList<IrRunFormat?> Formats, CellSpan? Cell)> Variants { get; } = new();
        public bool Used { get; set; }
    }

    /// <param name="fusedAnchors">Body paragraph anchors to leave out of pairing; null leaves none out.
    /// <see cref="ApplyToBoth"/> uses it to fall back to leaving fused paragraphs unpaired on both surfaces
    /// when the two scripts disagree on their move group ids.</param>
    public static IrEditScript Apply(
        IrEditScript script, IrDocument left, IrDocument right, IrDiffSettings settings,
        IReadOnlySet<string>? fusedAnchors = null) =>
        Pair(script, left, right, settings, fusedAnchors, view: null);

    /// <summary>
    /// Pair the relocations of a comparison whose redline fuses cross-paragraph runs (issue #930), once, for both
    /// of its scripts: <paramref name="data"/> (built without fusion; the revision list and edit script) and
    /// <paramref name="markup"/> (the fused build the redline draws). Pairing runs on the data script, where every
    /// paragraph has its own op. A half in a paragraph the markup script draws inside a cross-paragraph run is
    /// admitted only when the run holds it as one deleted (or inserted) stretch of one output paragraph — for a
    /// whole paragraph, an output paragraph of its own whose mark is deleted (or inserted) — and that stretch is
    /// tagged with the same id. Every other half must sit in an op the two scripts share value-for-value, whose
    /// rewrite is then carried over. A half either script could not draw as a move pairs on neither.
    /// </summary>
    public static (IrEditScript Data, IrEditScript Markup) ApplyToBoth(
        IrEditScript data, IrEditScript markup, IrDocument left, IrDocument right, IrDiffSettings settings)
    {
        if (!PairsAnything(settings))
            return (data, markup);
        // Relocation ids continue above the move groups; the scripts must agree on where that is.
        if (MaxMoveGroupId(data) != MaxMoveGroupId(markup))
        {
            var fused = new HashSet<string>(System.StringComparer.Ordinal);
            foreach (var op in markup.Operations.Where(o => o.Kind == IrEditOpKind.CrossParagraphRunBlock))
                foreach (var cell in op.CrossParagraphCells ?? IrNodeList.Empty<IrCrossParagraphCell>())
                {
                    if (cell.LeftAnchor is { } la) fused.Add(la);
                    if (cell.RightAnchor is { } ra) fused.Add(ra);
                }
            return (Apply(data, left, right, settings, fused), Apply(markup, left, right, settings));
        }
        var view = new RedlineView(markup, data, left, right, settings);
        var pairedData = Pair(data, left, right, settings, fusedAnchors: null, view);
        return (pairedData, view.Replay(data, pairedData));
    }

    /// <summary>Pairing only labels content as moved, so it runs only where a move is drawn and reported: move
    /// reporting on, the engine's fine revision grain (the WmlComparer-compatible grain reproduces that
    /// comparer's revision set, which has no such moves), and no input revisions being preserved.</summary>
    private static bool PairsAnything(IrDiffSettings settings) =>
        !settings.PreserveInputRevisions && settings.RenderMoves &&
        settings.RevisionGranularity == RevisionGranularity.Fine;

    private static IrEditScript Pair(
        IrEditScript script, IrDocument left, IrDocument right, IrDiffSettings settings,
        IReadOnlySet<string>? fusedAnchors, RedlineView? view)
    {
        if (!PairsAnything(settings))
            return script;

        int min = settings.MoveMinimumTokenCount;
        var blocks = new List<Candidate>();
        var hosts = new List<Host>();
        int nextList = 0;
        bool Excluded(string? anchor) => anchor is not null && fusedAnchors?.Contains(anchor) == true;
        // With a redline view, a top-level op is a candidate only when the markup script either has the same op
        // (its rewrite carries over) or draws its paragraphs inside a run (its halves are located in the cells);
        // a nested op inherits its table's standing.
        void Collect(IEnumerable<IrEditOp> ops, bool topLevel, bool inheritedShared)
        {
            int listId = nextList++;
            foreach (var op in ops)
            {
                bool fused = view is not null && topLevel && (view.IsFused(op.LeftAnchor) || view.IsFused(op.RightAnchor));
                bool shared = view is null || (topLevel ? view.Shares(op) : inheritedShared);
                if (fused || shared)
                {
                    var host = fused ? view : null;
                    if (op.Kind == IrEditOpKind.DeleteBlock && Paragraph(op.LeftAnchor, left) is { } lp)
                        AddBlock(op, lp, listId, isSource: true, host);
                    else if (op.Kind == IrEditOpKind.InsertBlock && Paragraph(op.RightAnchor, right) is { } rp)
                        AddBlock(op, rp, listId, isSource: false, host);
                    else if (op.Kind == IrEditOpKind.ModifyBlock && op.TokenDiff is { } diff &&
                             !op.RequiresWholeParagraphReplace && op.TextboxDiffs is null &&
                             Paragraph(op.LeftAnchor, left) is { } ml && Paragraph(op.RightAnchor, right) is { } mr)
                        hosts.Add(new Host
                        {
                            Op = op, ListId = listId, Atoms = Atomize(diff), FormatComparison = settings.FormatComparison,
                            Left = IrDiffTokenizer.Tokenize(ml, settings), Right = IrDiffTokenizer.Tokenize(mr, settings),
                            FusedIn = host,
                        });
                }
                foreach (var cellOps in CellOpLists(op, settings))
                    Collect(cellOps, topLevel: false, inheritedShared: shared && !fused);
            }
        }
        IrParagraph? Paragraph(string? anchor, IrDocument doc) =>
            anchor is not null && !Excluded(anchor) && doc.AnchorIndex.TryGetValue(anchor, out var block) &&
            block is IrParagraph paragraph ? paragraph : null;
        void AddBlock(IrEditOp op, IrParagraph paragraph, int listId, bool isSource, RedlineView? fusedIn)
        {
            var tokens = IrDiffTokenizer.Tokenize(paragraph, settings);
            var (text, words) = Normalize(tokens);
            if (words < min)
                return;
            CellSpan? cell = null;
            if (fusedIn is not null)
            {
                cell = fusedIn.LocateWhole(isSource ? op.LeftAnchor! : op.RightAnchor!, isSource, tokens.Count);
                if (cell is null)
                    return;
            }
            blocks.Add(new Candidate(op, listId, text, WordFormats(tokens, 0, tokens.Count), isSource, cell));
        }
        Collect(script.Operations, topLevel: true, inheritedShared: true);
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
            source.Cell?.Commit(nextGroup);
            destination.Cell?.Commit(nextGroup);
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
            span.Variants.Add((0, block.Text, block.Formats, block.Cell));
            spans.Add(span);
        }

        // Destinations by text; each source, in document order, claims an unused destination with the same
        // text in the same formatting. Preferred, in order: edges on a sentence end (so a move reads "The second
        // sentence." rather than "second sentence. The" when both halves can slide), the smallest total slide,
        // the earliest destination.
        var destinations = new Dictionary<string, List<(Span Span, int Offset, IReadOnlyList<IrRunFormat?> Formats, CellSpan? Cell)>>(
            System.StringComparer.Ordinal);
        foreach (var span in spans.Where(s => !s.IsSource))
            foreach (var (offset, text, formats, cell) in span.Variants)
            {
                if (!destinations.TryGetValue(text, out var list))
                    destinations[text] = list = new List<(Span, int, IReadOnlyList<IrRunFormat?>, CellSpan?)>();
                list.Add((span, offset, formats, cell));
            }

        foreach (var source in spans.Where(s => s.IsSource).OrderBy(s => s.Order))
        {
            (Span Span, int SourceOffset, int DestinationOffset, CellSpan? SourceCell, CellSpan? DestinationCell,
                (int Edge, int Slide, int Order) Key)? best = null;
            foreach (var (offset, text, formats, sourceCell) in source.Variants)
            {
                if (!destinations.TryGetValue(text, out var candidates))
                    continue;
                int edge = EndsSentence(text) ? 0 : 1;
                foreach (var (destination, destinationOffset, destinationFormats, destinationCell) in candidates)
                {
                    if (destination.Used || (source.Block != null && destination.Block != null) ||
                        !SameFormats(formats, destinationFormats, formatComparison) ||
                        !Separated(source, offset, destination, destinationOffset) ||
                        !CanClaim(source, offset) || !CanClaim(destination, destinationOffset) ||
                        !CellSpan.Compatible(sourceCell, destinationCell))
                        continue;
                    var key = (edge, System.Math.Abs(offset) + System.Math.Abs(destinationOffset), destination.Order);
                    if (best is null || key.CompareTo(best.Value.Key) < 0)
                        best = (destination, offset, destinationOffset, sourceCell, destinationCell, key);
                }
            }
            if (best is not { } chosen)
                continue;
            int id = nextGroup++;
            Commit(source, chosen.SourceOffset, id, groupOf);
            Commit(chosen.Span, chosen.DestinationOffset, id, groupOf);
            chosen.SourceCell?.Commit(id);
            chosen.DestinationCell?.Commit(id);
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
            if (words < min)
                return;
            CellSpan? cell = null;
            if (host.FusedIn is { } view)
            {
                cell = view.Locate(deleted ? host.Op.LeftAnchor! : host.Op.RightAnchor!, deleted,
                    s + offset, e + offset, tokens.Count);
                if (cell is null)
                    return;
            }
            span.Variants.Add((offset, text, WordFormats(tokens, s + offset, e + offset), cell));
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

    // ------------------------------------------------------------------ fused runs (issue #930)

    /// <summary>One output paragraph of a cross-paragraph run in the markup script, as atoms that pairing tags.</summary>
    private sealed class CellView
    {
        public CellView(IrCrossParagraphCell cell, IReadOnlyList<IrDiffToken> leftTokens)
        {
            Cell = cell;
            Atoms = Atomize(cell.Diff);
            LeftTokens = leftTokens;
        }

        public IrCrossParagraphCell Cell { get; }

        public List<Atom> Atoms { get; }

        /// <summary>The whole left paragraph's tokens (the cell's left slice starts at
        /// <see cref="IrCrossParagraphCell.LeftSliceStart"/>), empty when the cell has no left side.</summary>
        public IReadOnlyList<IrDiffToken> LeftTokens { get; }

        public HashSet<int> Claimed { get; } = new();

        public bool Changed { get; set; }
    }

    /// <summary>A relocation half located in a run's cell: the atoms [<see cref="AtomStart"/>,
    /// <see cref="AtomEnd"/>) of <see cref="Cell"/>, one contiguous deleted or inserted stretch.</summary>
    private sealed record CellSpan(CellView Cell, int AtomStart, int AtomEnd)
    {
        /// <summary>Whether two located halves can pair: each still unclaimed, and two halves in the same output
        /// paragraph separated by a retained word — adjacent, they would be drawn as a replacement.</summary>
        public static bool Compatible(CellSpan? source, CellSpan? destination)
        {
            if (source is not null && !source.Free())
                return false;
            if (destination is not null && !destination.Free())
                return false;
            if (source is null || destination is null || !ReferenceEquals(source.Cell, destination.Cell))
                return true;
            var cell = source.Cell;
            for (int i = System.Math.Min(source.AtomEnd, destination.AtomEnd);
                 i < System.Math.Max(source.AtomStart, destination.AtomStart); i++)
            {
                var atom = cell.Atoms[i];
                if (!IsAligned(atom))
                    continue;
                int index = cell.Cell.LeftSliceStart + atom.L;
                if (index >= cell.LeftTokens.Count)
                    return false; // the tokenizations disagree; fail closed
                if (cell.LeftTokens[index].Kind == IrDiffTokenKind.Word)
                    return true;
            }
            return false;
        }

        private bool Free()
        {
            for (int i = AtomStart; i < AtomEnd; i++)
                if (Cell.Claimed.Contains(i))
                    return false;
            return true;
        }

        public void Commit(int id)
        {
            for (int i = AtomStart; i < AtomEnd; i++)
            {
                Cell.Claimed.Add(i);
                Cell.Atoms[i] = Cell.Atoms[i] with { Relocation = id };
            }
            Cell.Changed = true;
        }
    }

    /// <summary>
    /// The markup script of a comparison whose redline fuses cross-paragraph runs, seen from its data script:
    /// which data ops it shares value-for-value, and where each fused paragraph's tokens sit in the run cells.
    /// </summary>
    private sealed class RedlineView
    {
        private readonly IrEditScript _markup;
        private readonly Dictionary<IrEditOp, IrEditOp> _markupOf = new(ReferenceEqualityComparer.Instance);
        private readonly Dictionary<IrEditOp, IrEditOp> _dataOf = new(ReferenceEqualityComparer.Instance);
        private readonly Dictionary<string, List<CellView>> _leftCells = new(System.StringComparer.Ordinal);
        private readonly Dictionary<string, List<CellView>> _rightCells = new(System.StringComparer.Ordinal);
        private readonly Dictionary<IrEditOp, List<CellView>> _runCells = new(ReferenceEqualityComparer.Instance);

        public RedlineView(IrEditScript markup, IrEditScript data, IrDocument left, IrDocument right, IrDiffSettings settings)
        {
            _markup = markup;
            var tokensOf = new Dictionary<string, IReadOnlyList<IrDiffToken>>(System.StringComparer.Ordinal);
            IReadOnlyList<IrDiffToken> Tokens(string anchor, IrDocument doc)
            {
                if (!tokensOf.TryGetValue(anchor, out var tokens))
                    tokensOf[anchor] = tokens = doc.AnchorIndex.TryGetValue(anchor, out var block) && block is IrParagraph p
                        ? IrDiffTokenizer.Tokenize(p, settings)
                        : System.Array.Empty<IrDiffToken>();
                return tokens;
            }

            var byKey = new Dictionary<(IrEditOpKind, string?, string?), IrEditOp?>();
            foreach (var op in markup.Operations)
            {
                if (op.Kind == IrEditOpKind.CrossParagraphRunBlock)
                {
                    var views = new List<CellView>();
                    foreach (var cell in op.CrossParagraphCells ?? IrNodeList.Empty<IrCrossParagraphCell>())
                    {
                        var view = new CellView(cell, cell.LeftAnchor is { } la ? Tokens(la, left) : System.Array.Empty<IrDiffToken>());
                        views.Add(view);
                        if (cell.LeftAnchor is { } l)
                            Add(_leftCells, l, view);
                        if (cell.RightAnchor is { } r)
                            Add(_rightCells, r, view);
                    }
                    _runCells[op] = views;
                    continue;
                }
                var key = (op.Kind, op.LeftAnchor, op.RightAnchor);
                byKey[key] = byKey.ContainsKey(key) ? null : op; // an ambiguous key shares nothing
            }
            foreach (var op in data.Operations)
                if (byKey.TryGetValue((op.Kind, op.LeftAnchor, op.RightAnchor), out var twin) && twin is not null &&
                    twin.Equals(op))
                {
                    _markupOf[op] = twin;
                    _dataOf[twin] = op;
                }

            static void Add(Dictionary<string, List<CellView>> index, string anchor, CellView view)
            {
                if (!index.TryGetValue(anchor, out var list))
                    index[anchor] = list = new List<CellView>();
                list.Add(view);
            }
        }

        public bool IsFused(string? anchor) =>
            anchor is not null && (_leftCells.ContainsKey(anchor) || _rightCells.ContainsKey(anchor));

        /// <summary>Whether the markup script holds this top-level data op unchanged.</summary>
        public bool Shares(IrEditOp dataOp) => _markupOf.ContainsKey(dataOp);

        /// <summary>
        /// Locate tokens [<paramref name="start"/>, <paramref name="end"/>) of a fused paragraph's
        /// <paramref name="leftSide"/> in the run: they must lie in one cell, as one contiguous stretch of deleted
        /// (left) or inserted (right) atoms. Null when they do not, or when a cell's slice reaches past the
        /// paragraph's <paramref name="tokenCount"/> tokens (the two tokenizations disagree; fail closed).
        /// </summary>
        public CellSpan? Locate(string anchor, bool leftSide, int start, int end, int tokenCount)
        {
            if (!(leftSide ? _leftCells : _rightCells).TryGetValue(anchor, out var cells))
                return null;
            CellView? home = null;
            foreach (var cell in cells)
            {
                int sliceStart = leftSide ? cell.Cell.LeftSliceStart : cell.Cell.RightSliceStart;
                int sliceEnd = sliceStart + (leftSide ? cell.Cell.LeftSliceLen : cell.Cell.RightSliceLen);
                if (sliceEnd > tokenCount)
                    return null;
                if (start >= sliceStart && end <= sliceEnd)
                    home = cell;
            }
            if (home is null)
                return null;

            int offset = leftSide ? home.Cell.LeftSliceStart : home.Cell.RightSliceStart;
            var kind = leftSide ? IrTokenOpKind.Delete : IrTokenOpKind.Insert;
            int first = -1, last = -1;
            for (int i = 0; i < home.Atoms.Count; i++)
            {
                int index = leftSide ? home.Atoms[i].L : home.Atoms[i].R;
                if (index < 0 || index + offset < start || index + offset >= end)
                    continue;
                if (home.Atoms[i].Kind != kind || (last >= 0 && i != last + 1))
                    return null;
                if (first < 0)
                    first = i;
                last = i;
            }
            return first >= 0 && last - first + 1 == end - start ? new CellSpan(home, first, last + 1) : null;
        }

        /// <summary>Locate a whole fused paragraph: all its tokens in one output paragraph that is its own and is
        /// removed (a left paragraph, mark deleted) or introduced (a right paragraph, mark inserted).</summary>
        public CellSpan? LocateWhole(string anchor, bool leftSide, int tokenCount)
        {
            var located = Locate(anchor, leftSide, 0, tokenCount, tokenCount);
            if (located is null)
                return null;
            var cell = located.Cell.Cell;
            bool own = leftSide
                ? cell.Mark == IrCrossParagraphMark.Deleted && cell.LeftSliceStart == 0 && cell.LeftSliceLen == tokenCount
                : cell.Mark == IrCrossParagraphMark.Inserted && cell.RightSliceStart == 0 && cell.RightSliceLen == tokenCount;
            return own ? located : null;
        }

        /// <summary>The markup script with the pairing carried over: each shared op replaced by its rewritten data
        /// twin, and each run whose cells were tagged rebuilt from their atoms.</summary>
        public IrEditScript Replay(IrEditScript data, IrEditScript pairedData)
        {
            var rewritten = new Dictionary<IrEditOp, IrEditOp>(ReferenceEqualityComparer.Instance);
            for (int i = 0; i < data.Operations.Count; i++)
                if (!ReferenceEquals(data.Operations[i], pairedData.Operations[i]))
                    rewritten[data.Operations[i]] = pairedData.Operations[i];
            if (rewritten.Count == 0 && !_runCells.Values.Any(cells => cells.Any(c => c.Changed)))
                return _markup;
            return _markup with
            {
                Operations = IrNodeList.From(_markup.Operations.Select(op =>
                {
                    if (_runCells.TryGetValue(op, out var cells) && cells.Any(c => c.Changed))
                        return op with
                        {
                            CrossParagraphCells = IrNodeList.From(cells.Select(c =>
                                c.Changed ? c.Cell with { Diff = new IrTokenDiff(Regroup(c.Atoms)) } : c.Cell)),
                        };
                    return _dataOf.TryGetValue(op, out var twin) && rewritten.TryGetValue(twin, out var r) ? r : op;
                })),
            };
        }
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
