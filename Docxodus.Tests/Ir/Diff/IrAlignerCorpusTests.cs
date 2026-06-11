#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Docxodus;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;
using Xunit.Abstractions;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// M2.1 Task 3 corpus smoke: run <see cref="IrBlockAligner"/> over every base↔variant DOCX pair we can
/// infer from <c>TestFiles/WC/</c> by name convention, asserting totality (no throw) + the shared
/// invariants both forward (before→after) and reversed (after→before), and logging a per-pair entry-kind
/// histogram plus corpus totals.
/// </summary>
/// <remarks>
/// <para><b>Pairing rules (inferred by name inspection; documented so the pair list is reproducible).</b>
/// A WC file name is <c>WCnnn-Name[-suffix].docx</c> (or the prefix-less <c>WC-BodyBookmarks-…</c>).
/// We group files by their <em>family key</em> = the longest leading segment that names a base document,
/// then pair the family's BASE against every VARIANT in the family:</para>
/// <list type="number">
/// <item><b>-Before/-After family.</b> If a family has a <c>…-Before</c> file, it is the base; every
/// <c>…-After</c>, <c>…-After1</c>, <c>…-After2</c>, <c>…-After-1</c>, … is a variant. The family key is the
/// name with the trailing <c>-Before</c>/<c>-After[n]</c>/<c>-After-n</c> token stripped (e.g.
/// <c>WC011-Before</c> + <c>WC011-After</c>; <c>WC013-Image-Before</c> + <c>WC013-Image-After</c> +
/// <c>WC013-Image-After2</c>; the <c>…-Before2</c>/<c>…-After2</c> pair forms its own family).</item>
/// <item><b>Base ↔ -Mod / -Deleted-* / other-variant family.</b> Otherwise the base is the file with NO
/// extra trailing edit token — the shortest name in the family (e.g. <c>WC001-Digits</c>,
/// <c>WC002-Unmodified</c>, <c>WC006-Table</c>) — and every longer sibling sharing that base as a prefix
/// is a variant (e.g. <c>WC001-Digits</c> ↔ <c>WC001-Digits-Mod</c> AND ↔
/// <c>WC001-Digits-Deleted-Paragraph</c>; <c>WC006-Table</c> ↔ <c>WC006-Table-Delete-Row</c> ↔
/// <c>WC006-Table-Delete-Contests-of-Row</c>).</item>
/// <item><b>WCnnn-prefix fan-out.</b> For families where no single shortest base prefixes the others
/// (e.g. <c>WC002-DeleteAtBeginning</c>, <c>WC002-InsertAtEnd</c>, … all share only the <c>WC002-</c>
/// numeric prefix), the designated base is the <c>…-Unmodified</c> sibling when present, paired against
/// every other <c>WCnnn-</c> sibling. <c>WC002</c> has <c>WC002-Unmodified</c>; <c>WC007</c> has
/// <c>WC007-Unmodified</c>.</item>
/// </list>
/// <para>The concrete pair list is built by <see cref="BuildPairs"/> below; it is logged so a reviewer
/// can see exactly what ran (92 pairs covering 161 of 163 WC files). Every pair is exercised forward
/// AND reversed. Two files are deliberately left unpaired — <c>WC014-SmartArt-With-Image-Deleted-After</c>
/// and <c>…-Deleted-After2</c> — because the <c>-Deleted-</c> infix gives them no unambiguous base under
/// the rules above (their natural base would be the sibling <c>…-Before</c>, but pairing across a
/// <c>-Deleted-</c> infix is left out rather than special-cased); the alignment of that family's
/// straightforward <c>-Before</c>/<c>-After</c>/<c>-After2</c> pairs already exercises the same content.</para>
/// <para>IR is read with <c>RetainSources = false</c> and <c>RevisionView.Accept</c> (the aligner needs
/// only reader-computed hashes; accepted revisions give a clean post-edit block stream).</para>
/// </remarks>
[Trait("Category", "Corpus")]
public class IrAlignerCorpusTests
{
    private static readonly IrReaderOptions ReadOpts =
        new() { RetainSources = false, RevisionView = RevisionView.Accept };
    private static readonly IrDiffSettings Diff = new();

    private static DirectoryInfo WcDir => new(Path.Combine("../../../../TestFiles/WC"));

    private readonly ITestOutputHelper _out;

    public IrAlignerCorpusTests(ITestOutputHelper output) => _out = output;

    [Fact]
    public void WC_corpus_pairs_align_without_throwing_invariants_hold_both_directions()
    {
        var pairs = BuildPairs();
        Assert.True(pairs.Count >= 30,
            $"Expected a substantial WC pair list; only inferred {pairs.Count}. Naming convention drift?");

        var totals = new Dictionary<IrAlignmentKind, int>();
        foreach (var k in Enum.GetValues<IrAlignmentKind>())
            totals[k] = 0;

        _out.WriteLine($"WC corpus: {pairs.Count} base↔variant pairs (each run forward + reversed)");
        _out.WriteLine("");

        foreach (var (baseName, variantName) in pairs)
        {
            var baseDoc = ReadWc(baseName);
            var variantDoc = ReadWc(variantName);

            // Forward: before → after.
            var fwd = IrBlockAligner.Align(baseDoc, variantDoc, Diff);
            IrAlignmentAsserts.AssertInvariants(baseDoc, variantDoc, fwd);

            // Reversed: after → before. Invariants must hold both directions.
            var rev = IrBlockAligner.Align(variantDoc, baseDoc, Diff);
            IrAlignmentAsserts.AssertInvariants(variantDoc, baseDoc, rev);

            foreach (var e in fwd.Entries)
                totals[e.Kind]++;

            _out.WriteLine($"  {baseName} -> {variantName}");
            _out.WriteLine($"      fwd: {IrAlignmentAsserts.Histogram(fwd)}");
            _out.WriteLine($"      rev: {IrAlignmentAsserts.Histogram(rev)}");
        }

        _out.WriteLine("");
        _out.WriteLine("Corpus totals (forward direction):");
        foreach (var kv in totals.OrderBy(k => (int)k.Key))
            _out.WriteLine($"  {kv.Key} = {kv.Value}");
    }

    private static IrDocument ReadWc(string fileName)
    {
        var fi = new FileInfo(Path.Combine(WcDir.FullName, fileName));
        Assert.True(fi.Exists, $"Missing WC test file: {fi.FullName}");
        return IrReader.Read(new WmlDocument(fi.FullName), ReadOpts);
    }

    // ------------------------------------------------------------------ pairing

    /// <summary>
    /// Build the (base, variant) file-name pair list from the WC directory by the documented rules.
    /// Deterministic: families and variants are sorted by ordinal name.
    /// </summary>
    private static List<(string Base, string Variant)> BuildPairs()
    {
        var files = WcDir.GetFiles("*.docx")
            .Select(f => f.Name)
            .OrderBy(n => n, StringComparer.Ordinal)
            .ToList();

        var pairs = new List<(string, string)>();
        var consumed = new HashSet<string>(StringComparer.Ordinal);

        // --- Rule 1: -Before… / -After… families.
        // Split each name at the FIRST "-Before" or "-After" token; the prefix is the family key and
        // the token+tail (e.g. "", "2", "-1", "-Delete-1-Row") is the variant index. A "Before" file is
        // a base, an "After" file is a variant. When a family has index-matched befores AND afters
        // (the WC021 "Before-1/After-1", "Before-2/After-2" case) we pair by matching index; otherwise
        // the single before pairs against EVERY after (the WC033/WC034 "Before" ↔ "After1/2/3" case).
        var families = files
            .Select(n => (Name: n, Split: SplitBeforeAfter(n)))
            .Where(t => t.Split is not null)
            .GroupBy(t => t.Split!.Value.Family, StringComparer.Ordinal);

        foreach (var g in families.OrderBy(x => x.Key, StringComparer.Ordinal))
        {
            var befores = g.Where(t => t.Split!.Value.IsBefore)
                .Select(t => (t.Name, t.Split!.Value.Index))
                .OrderBy(t => t.Name, StringComparer.Ordinal).ToList();
            var afters = g.Where(t => !t.Split!.Value.IsBefore)
                .Select(t => (t.Name, t.Split!.Value.Index))
                .OrderBy(t => t.Name, StringComparer.Ordinal).ToList();
            if (befores.Count == 0 || afters.Count == 0)
                continue; // an -After family with no -Before (or vice-versa): leave for later rules

            if (befores.Count > 1)
            {
                // Multiple bases: pair each before with the after(s) sharing its index; un-indexed
                // afters fall back to the lexically-first before.
                foreach (var (afterName, afterIdx) in afters)
                {
                    var (beforeName, _) = befores.FirstOrDefault(b => b.Index == afterIdx);
                    beforeName ??= befores[0].Name;
                    AddPair(pairs, consumed, beforeName, afterName);
                }
            }
            else
            {
                foreach (var (afterName, _) in afters)
                    AddPair(pairs, consumed, befores[0].Name, afterName);
            }
        }

        // --- Rule 2: base ↔ prefix-extending variant families (Mod, Deleted-*, etc).
        // Among files NOT already consumed, a base is one that is a strict name-prefix (at a '-'
        // boundary) of one or more siblings; pair it with every sibling it prefixes.
        var remaining = files.Where(n => !consumed.Contains(n)).ToList();
        foreach (var baseFile in remaining.OrderBy(n => n, StringComparer.Ordinal))
        {
            string baseStem = Stem(baseFile);
            var variants = remaining
                .Where(other => !ReferenceEquals(other, baseFile) && other != baseFile)
                .Where(other => IsVariantOf(baseStem, other))
                .OrderBy(n => n, StringComparer.Ordinal)
                .ToList();
            foreach (var variant in variants)
                AddPair(pairs, consumed, baseFile, variant);
        }

        // --- Rule 3: WCnnn-prefix fan-out around an -Unmodified base (e.g. WC002, WC007).
        // Group leftover files by their WCnnn numeric prefix; if the group has an -Unmodified file,
        // pair it against every other sibling in the group.
        var leftover = files.Where(n => !consumed.Contains(n)).ToList();
        var byNum = leftover
            .Select(n => (Name: n, Num: NumericPrefix(n)))
            .Where(t => t.Num is not null)
            .GroupBy(t => t.Num!, StringComparer.Ordinal);

        foreach (var g in byNum.OrderBy(g => g.Key, StringComparer.Ordinal))
        {
            var members = g.Select(t => t.Name).OrderBy(n => n, StringComparer.Ordinal).ToList();
            var baseFile = members.FirstOrDefault(m =>
                Stem(m).EndsWith("Unmodified", StringComparison.Ordinal));
            if (baseFile is null)
                continue;
            foreach (var variant in members.Where(m => m != baseFile))
                AddPair(pairs, consumed, baseFile, variant);
        }

        return pairs
            .Distinct()
            .OrderBy(p => p.Item1, StringComparer.Ordinal)
            .ThenBy(p => p.Item2, StringComparer.Ordinal)
            .ToList();
    }

    private static string Stem(string fileName) =>
        fileName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)
            ? fileName[..^5]
            : fileName;

    private static void AddPair(
        List<(string, string)> pairs, HashSet<string> consumed, string baseFile, string variant)
    {
        pairs.Add((baseFile, variant));
        consumed.Add(baseFile);
        consumed.Add(variant);
    }

    /// <summary>
    /// Split a name at its FIRST <c>-Before</c>/<c>-After</c> token into (family prefix, isBefore,
    /// index = the token tail, e.g. "" / "2" / "-1" / "-Delete-1-Row"). Returns null when the name has
    /// no such token. The family prefix is shared by a base and all its before/after variants; the
    /// index distinguishes index-matched before/after pairs within a multi-base family.
    /// </summary>
    private static (string Family, bool IsBefore, string Index)? SplitBeforeAfter(string fileName)
    {
        string stem = Stem(fileName);
        int b = stem.IndexOf("-Before", StringComparison.Ordinal);
        int a = stem.IndexOf("-After", StringComparison.Ordinal);
        if (b < 0 && a < 0)
            return null;

        bool isBefore = b >= 0 && (a < 0 || b < a);
        int idx = isBefore ? b : a;
        string token = isBefore ? "-Before" : "-After";
        string family = stem[..idx];
        string indexTail = stem[(idx + token.Length)..]; // "", "2", "-1", "-Delete-1-Row", …
        return (family, isBefore, indexTail);
    }

    /// <summary>True if <paramref name="other"/> extends <paramref name="baseStem"/> at a '-' boundary.</summary>
    private static bool IsVariantOf(string baseStem, string other) =>
        Stem(other).Length > baseStem.Length &&
        Stem(other).StartsWith(baseStem + "-", StringComparison.Ordinal);

    /// <summary>The leading <c>WCnnn</c> numeric token, or null for non-conforming names.</summary>
    private static string? NumericPrefix(string fileName)
    {
        string stem = Stem(fileName);
        if (!stem.StartsWith("WC", StringComparison.Ordinal))
            return null;
        int i = 2;
        while (i < stem.Length && char.IsDigit(stem[i]))
            i++;
        return i > 2 ? stem[..i] : null;
    }
}
