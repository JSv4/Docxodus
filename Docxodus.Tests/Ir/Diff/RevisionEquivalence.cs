#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Docxodus;
using Docxodus.Ir.Diff;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// The SHARED semantic-equivalence core for comparing the two revision engines' output, extracted from the
/// M2.3 Task 2 differential harness (<see cref="IrVsWmlComparerTests"/>) so the Task 3 generative fuzzer
/// (<see cref="IrDiffFuzzTests"/>) reuses the IDENTICAL normalization + combined char-bag contract rather
/// than re-deriving it. The Task 2 harness keeps its richer cause-bucketing private; this class holds the
/// two primitives both sites need:
/// <list type="number">
/// <item><see cref="Normalize"/> — the precise text normalization (collapse whitespace runs to one space,
/// trim, drop-if-empty; CASE PRESERVED).</item>
/// <item><see cref="RevisionBag"/> — per-kind normalized-text multisets plus the granularity-independent
/// COMBINED Inserted+Deleted char bag, with <see cref="RevisionBag.MultisetsEqual"/> /
/// <see cref="RevisionBag.IsCombinedCharBagEquivalent"/> the two equivalence relations the fuzzer's
/// differential gate rides on.</item>
/// </list>
/// </summary>
internal static class RevisionEquivalence
{
    /// <summary>
    /// The comparison normalization, applied identically to both engines' revision text: every run of
    /// whitespace collapses to a single ASCII space, the result is trimmed, an empty result yields the
    /// empty string (callers drop empties). <b>Case is preserved</b> — a case change is a real content edit.
    /// This is byte-for-byte the Task 2 normalization contract.
    /// </summary>
    public static string Normalize(string? text)
    {
        if (string.IsNullOrEmpty(text))
            return string.Empty;
        var sb = new StringBuilder(text!.Length);
        bool pendingSpace = false;
        bool sawNonSpace = false;
        foreach (char c in text)
        {
            if (char.IsWhiteSpace(c))
            {
                pendingSpace = sawNonSpace; // suppress leading whitespace
                continue;
            }
            if (pendingSpace)
            {
                sb.Append(' ');
                pendingSpace = false;
            }
            sb.Append(c);
            sawNonSpace = true;
        }
        return sb.ToString();
    }

    /// <summary>
    /// Per-kind multisets of normalized revision text plus the combined Inserted+Deleted char bag. Both
    /// engines map into this shape via <see cref="FromIr"/> / <see cref="FromWmlComparer"/>; equivalence is
    /// pure set algebra over it.
    /// </summary>
    public sealed class RevisionBag
    {
        private readonly Dictionary<IrRevisionType, Dictionary<string, int>> _byKind = new();

        public int Total { get; private set; }

        private RevisionBag()
        {
            foreach (var k in Enum.GetValues<IrRevisionType>())
                _byKind[k] = new Dictionary<string, int>(StringComparer.Ordinal);
        }

        public static RevisionBag FromIr(IEnumerable<IrRevision> revs)
        {
            var b = new RevisionBag();
            foreach (var r in revs)
                b.Add(r.Type, r.Text);
            return b;
        }

        public static RevisionBag FromWmlComparer(IEnumerable<WmlComparer.WmlComparerRevision> revs)
        {
            var b = new RevisionBag();
            foreach (var r in revs)
                b.Add(MapKind(r.RevisionType), r.Text);
            return b;
        }

        private static IrRevisionType MapKind(WmlComparer.WmlComparerRevisionType t) => t switch
        {
            WmlComparer.WmlComparerRevisionType.Inserted => IrRevisionType.Inserted,
            WmlComparer.WmlComparerRevisionType.Deleted => IrRevisionType.Deleted,
            WmlComparer.WmlComparerRevisionType.Moved => IrRevisionType.Moved,
            WmlComparer.WmlComparerRevisionType.FormatChanged => IrRevisionType.FormatChanged,
            _ => throw new ArgumentOutOfRangeException(nameof(t), t, "Unknown WmlComparerRevisionType"),
        };

        private void Add(IrRevisionType kind, string? rawText)
        {
            string norm = Normalize(rawText);
            if (norm.Length == 0)
                return; // empty-after-normalization atoms are dropped on both sides
            var d = _byKind[kind];
            d[norm] = d.TryGetValue(norm, out var n) ? n + 1 : 1;
            Total++;
        }

        public IReadOnlyDictionary<string, int> Multiset(IrRevisionType kind) => _byKind[kind];

        /// <summary>Exact per-kind normalized-text multiset equality across all four kinds.</summary>
        public bool MultisetsEqual(RevisionBag other) =>
            Enum.GetValues<IrRevisionType>().All(k => MultisetEqual(_byKind[k], other._byKind[k]));

        /// <summary>
        /// The loose, granularity-independent equivalence: the WHITESPACE-FREE combined Inserted+Deleted
        /// char bags are equal AND the Moved + FormatChanged multisets match exactly. Captures "same
        /// content, different atomization" — one engine reporting <c>"ab"</c> as one revision and the other
        /// <c>"a","b"</c> as two compare equal — while holding moves/format exact (the fuzzer excludes those
        /// kinds from cross-engine comparison anyway, so requiring them empty/equal here is conservative).
        /// </summary>
        public bool IsCombinedCharBagEquivalent(RevisionBag other) =>
            CharBagsEqual(InsDelCharBag(), other.InsDelCharBag()) &&
            MultisetEqual(_byKind[IrRevisionType.Moved], other._byKind[IrRevisionType.Moved]) &&
            MultisetEqual(_byKind[IrRevisionType.FormatChanged], other._byKind[IrRevisionType.FormatChanged]);

        /// <summary>
        /// WHITESPACE-FREE char multiset across all Inserted + Deleted atoms (granularity-independent):
        /// inter-word boundary spaces are an atomization artifact, not content, so they are excluded.
        /// </summary>
        public Dictionary<char, int> InsDelCharBag()
        {
            var bag = new Dictionary<char, int>();
            AddCharsOf(_byKind[IrRevisionType.Inserted], bag);
            AddCharsOf(_byKind[IrRevisionType.Deleted], bag);
            return bag;
        }

        private static void AddCharsOf(Dictionary<string, int> multiset, Dictionary<char, int> bag)
        {
            foreach (var (text, count) in multiset)
                for (int i = 0; i < count; i++)
                    foreach (char c in text)
                    {
                        if (char.IsWhiteSpace(c))
                            continue;
                        bag[c] = bag.TryGetValue(c, out var n) ? n + 1 : 1;
                    }
        }

        /// <summary>A deterministic dump of one kind's multiset, ordinal-sorted, for failure artifacts.</summary>
        public string Dump(IrRevisionType kind)
        {
            var d = _byKind[kind];
            if (d.Count == 0)
                return "(none)";
            return string.Join(", ", d.OrderBy(kv => kv.Key, StringComparer.Ordinal)
                .Select(kv => kv.Value == 1 ? $"\"{kv.Key}\"" : $"\"{kv.Key}\"×{kv.Value}"));
        }
    }

    public static bool MultisetEqual(IReadOnlyDictionary<string, int> a, IReadOnlyDictionary<string, int> b)
    {
        if (a.Count != b.Count)
            return false;
        foreach (var (k, v) in a)
            if (!b.TryGetValue(k, out var w) || w != v)
                return false;
        return true;
    }

    public static bool CharBagsEqual(Dictionary<char, int> a, Dictionary<char, int> b)
    {
        if (a.Count != b.Count)
            return false;
        foreach (var (k, v) in a)
            if (!b.TryGetValue(k, out var w) || w != v)
                return false;
        return true;
    }
}
