#nullable enable

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;
using Xunit.Abstractions;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// The frozen WC-corpus revision baseline — DocxDiff's standing regression net over the 92-pair
/// WmlComparer corpus, both directions, in both revision granularities.
///
/// <para><b>Why this file exists.</b> Until v11.0.0 this corpus was checked DIFFERENTIALLY: the
/// removed <c>IrVsWmlComparerTests</c> ran the legacy engine and the IR pipeline head to head and
/// classified every (pair, direction) by semantic agreement of their revision sets. That harness
/// could not survive the legacy engine's removal, so the last differential run's verdict was frozen
/// instead. Its four scoreboard reports are committed verbatim under
/// <c>docs/architecture/wmlcomparer_parity_baseline/</c> as the provenance record; this test is the
/// live half — it pins the IR pipeline's own output for every corpus pair so a regression still
/// fails a build even though the oracle that originally blessed these numbers is gone.</para>
///
/// <para><b>What is pinned.</b> For each (pair, direction, granularity) the per-kind multiset of
/// normalized revision text — the same <c>(kind, normtext)</c> atoms and the same normalization the
/// differential harness compared on (whitespace runs collapse to one space, ends trimmed, atoms
/// empty after normalization dropped, case preserved). Pinning the multiset rather than a count
/// means a regression that swaps WHICH text is inserted, while keeping the tally, still fails.</para>
///
/// <para><b>Header/footer scope is included here.</b> The differential harness filtered header- and
/// footer-scoped revisions out of the Fine-mode comparison because the legacy oracle structurally
/// could not report them — an exclusion that existed only to keep the two engines comparable. With
/// no oracle left the exclusion has no purpose, so the baseline pins the full revision set.</para>
///
/// <para><b>Regenerating.</b> A deliberate engine change moves these numbers. Set
/// <c>DOCXODUS_REGEN_CORPUS_BASELINE=1</c> and run this test: it rewrites the .tsv in place and
/// fails with a notice so the rewrite can never be mistaken for a pass. Review the resulting diff
/// as carefully as you would review the engine change — it IS the engine change, made visible.</para>
/// </summary>
public class DocxDiffCorpusBaselineTests
{
    private readonly ITestOutputHelper _out;

    public DocxDiffCorpusBaselineTests(ITestOutputHelper output) => _out = output;

    /// <summary>Fine granularity — the engine-native one-revision-per-token-span grain.</summary>
    private static readonly IrDiffSettings FineDiff = new();

    /// <summary>The coarser contiguous-region grain the legacy comparer reported.</summary>
    private static readonly IrDiffSettings CompatDiff =
        new() { RevisionGranularity = RevisionGranularity.WmlComparerCompatible };

    /// <summary>Committed next to this test so a reviewer sees the baseline move in the same diff.</summary>
    private static string BaselinePath =>
        Path.Combine("../../../Ir/Diff", "DocxDiffCorpusBaseline.tsv");

    private const string RegenVar = "DOCXODUS_REGEN_CORPUS_BASELINE";

    [Fact]
    public void DocxDiff_reproduces_the_frozen_WC_corpus_baseline()
    {
        var pairs = WcCorpus.BuildPairs();
        Assert.True(pairs.Count >= 30, $"Expected a substantial WC pair list; got {pairs.Count}.");

        var actual = new SortedDictionary<string, string>(StringComparer.Ordinal);
        foreach (var (baseName, variantName) in pairs)
        {
            Record(actual, baseName, variantName);
            Record(actual, variantName, baseName);
        }

        if (Environment.GetEnvironmentVariable(RegenVar) == "1")
        {
            Write(BaselinePath, actual);
            Assert.Fail(
                $"Baseline REGENERATED at {Path.GetFullPath(BaselinePath)} ({actual.Count} rows). " +
                $"Unset {RegenVar} and re-run to verify, and review the .tsv diff as an engine change.");
        }

        Assert.True(
            File.Exists(BaselinePath),
            $"Missing corpus baseline at {Path.GetFullPath(BaselinePath)}. " +
            $"Regenerate with {RegenVar}=1 only if it was deliberately deleted.");

        var expected = Read(BaselinePath);

        // Report before asserting so a failing run shows the whole picture, not just the first diff.
        var missing = expected.Keys.Where(k => !actual.ContainsKey(k)).ToList();
        var added = actual.Keys.Where(k => !expected.ContainsKey(k)).ToList();
        var changed = expected.Keys
            .Where(k => actual.TryGetValue(k, out var a) && !string.Equals(a, expected[k], StringComparison.Ordinal))
            .ToList();

        _out.WriteLine("===== DOCXDIFF WC-CORPUS BASELINE =====");
        _out.WriteLine($"Pairs: {pairs.Count} × 2 directions × 2 granularities = {actual.Count} rows");
        _out.WriteLine($"Matched: {expected.Count - missing.Count - changed.Count}   " +
                       $"Changed: {changed.Count}   Missing: {missing.Count}   Added: {added.Count}");

        foreach (var k in changed.Take(20))
            _out.WriteLine($"  CHANGED {k}\n    expected: {expected[k]}\n    actual:   {actual[k]}");
        foreach (var k in missing.Take(20))
            _out.WriteLine($"  MISSING  {k}  (baseline had: {expected[k]})");
        foreach (var k in added.Take(20))
            _out.WriteLine($"  ADDED    {k}  (now produces: {actual[k]})");

        Assert.True(
            missing.Count == 0 && added.Count == 0 && changed.Count == 0,
            $"DocxDiff no longer reproduces the frozen WC-corpus baseline: " +
            $"{changed.Count} changed, {missing.Count} missing, {added.Count} added. " +
            $"See the test output for the per-row diff. If the change is intended, regenerate with " +
            $"{RegenVar}=1 and review the .tsv diff.");
    }

    /// <summary>
    /// Render one direction of one pair in both granularities and fold each into the row map.
    /// A throw is a regression, not a recorded outcome — the IR pipeline must handle every corpus pair.
    /// </summary>
    private static void Record(
        SortedDictionary<string, string> rows, string leftName, string rightName)
    {
        string label = $"{Stem(leftName)} -> {Stem(rightName)}";

        var irLeft = WcCorpus.ReadWc(leftName);
        var irRight = WcCorpus.ReadWc(rightName);
        var script = IrEditScriptBuilder.Build(irLeft, irRight, FineDiff);

        rows[$"{label}\tfine"] =
            Encode(IrRevisionRenderer.Render(script, irLeft, irRight, FineDiff));
        rows[$"{label}\tcompat"] =
            Encode(IrRevisionRenderer.Render(script, irLeft, irRight, CompatDiff));
    }

    /// <summary>
    /// Serialize a revision set as its per-kind <c>(kind, normtext)</c> multiset, deterministically:
    /// kind in enum order, then normtext ordinal-ascending, each as <c>Kind|count|text</c>. Atoms empty
    /// after normalization are dropped (they carry no comparable content), so a renderer emitting a
    /// text-free marker revision does not perturb the baseline.
    /// </summary>
    private static string Encode(IEnumerable<IrRevision> revisions)
    {
        var byKind = new Dictionary<IrRevisionType, Dictionary<string, int>>();
        foreach (var kind in Enum.GetValues<IrRevisionType>())
            byKind[kind] = new Dictionary<string, int>(StringComparer.Ordinal);

        foreach (var r in revisions)
        {
            string norm = Normalize(r.Text);
            if (norm.Length == 0)
                continue;
            var bag = byKind[r.Type];
            bag[norm] = bag.TryGetValue(norm, out int n) ? n + 1 : 1;
        }

        var parts = new List<string>();
        foreach (var kind in Enum.GetValues<IrRevisionType>())
        {
            foreach (var entry in byKind[kind].OrderBy(e => e.Key, StringComparer.Ordinal))
            {
                parts.Add(string.Create(
                    CultureInfo.InvariantCulture,
                    $"{kind}|{entry.Value}|{Escape(entry.Key)}"));
            }
        }

        if (parts.Count == 0)
            return "(none)";

        string full = string.Join("  ", parts);
        if (full.Length <= MaxInlineRowLength)
            return full;

        // A handful of corpus pairs rewrite most of the document, so their literal multiset runs to
        // hundreds of KB and would dominate the committed baseline (the median row is under 100 bytes).
        // Pin those by digest instead: a SHA-256 over the SAME deterministic encoding still fails on any
        // change, and the per-kind tallies keep the row diagnosable without storing the text.
        var tally = string.Join(
            ",",
            Enum.GetValues<IrRevisionType>()
                .Select(k => $"{k}={byKind[k].Values.Sum()}")
                .Where(s => !s.EndsWith("=0", StringComparison.Ordinal)));

        return string.Create(
            CultureInfo.InvariantCulture,
            $"#digest atoms[{tally}] sha256:{Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(full))).ToLowerInvariant()}");
    }

    /// <summary>
    /// Rows longer than this are pinned by digest rather than stored literally. Chosen so every
    /// ordinary corpus row stays fully readable in review and only whole-document rewrites collapse.
    /// </summary>
    private const int MaxInlineRowLength = 2000;

    /// <summary>
    /// Collapse every whitespace run to a single space and trim both ends; case is preserved (a case
    /// flip is a real content change). Mirrors the normalization the differential harness compared on,
    /// so a baseline row stays comparable to that harness's recorded atoms.
    /// </summary>
    private static string Normalize(string? text)
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

    /// <summary>Keep one row on one line and the column separator unambiguous.</summary>
    private static string Escape(string s) =>
        s.Replace("\\", "\\\\", StringComparison.Ordinal)
         .Replace("\t", "\\t", StringComparison.Ordinal)
         .Replace("\n", "\\n", StringComparison.Ordinal)
         .Replace("\r", "\\r", StringComparison.Ordinal);

    private static string Stem(string fileName) => Path.GetFileNameWithoutExtension(fileName);

    private static void Write(string path, SortedDictionary<string, string> rows)
    {
        var sb = new StringBuilder();
        sb.Append("# DocxDiff WC-corpus revision baseline. Columns: label <TAB> granularity <TAB> revisions.\n");
        sb.Append("# Revisions are the per-kind (kind, normalized-text) multiset: Kind|count|text, joined by two spaces.\n");
        sb.Append($"# Regenerate with {RegenVar}=1; review the diff as an engine change.\n");
        foreach (var kv in rows)
            sb.Append(kv.Key).Append('\t').Append(kv.Value).Append('\n');

        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
        File.WriteAllText(path, sb.ToString());
    }

    private static SortedDictionary<string, string> Read(string path)
    {
        var rows = new SortedDictionary<string, string>(StringComparer.Ordinal);
        foreach (var line in File.ReadAllLines(path))
        {
            if (line.Length == 0 || line[0] == '#')
                continue;
            int last = line.LastIndexOf('\t');
            Assert.True(last > 0, $"Malformed baseline row: {line}");
            rows[line[..last]] = line[(last + 1)..];
        }
        return rows;
    }
}
