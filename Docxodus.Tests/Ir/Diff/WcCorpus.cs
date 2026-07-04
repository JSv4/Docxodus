#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Docxodus;
using Docxodus.Ir;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Shared WC-corpus pairing + reading infrastructure for the diff corpus tests
/// (<see cref="IrAlignerCorpusTests"/> and <see cref="IrEditScriptCorpusTests"/>). The pair list is
/// inferred from <c>TestFiles/WC/</c> by name convention — see <see cref="BuildPairs"/> for the
/// documented rules; it yields 92 base↔variant pairs covering 161 of 163 WC files.
/// </summary>
internal static class WcCorpus
{
    public static readonly IrReaderOptions ReadOpts =
        new() { RetainSources = false, RevisionView = RevisionView.Accept };

    public static DirectoryInfo WcDir => new(Path.Combine("../../../../TestFiles/WC"));

    public static IrDocument ReadWc(string fileName)
    {
        var fi = new FileInfo(Path.Combine(WcDir.FullName, fileName));
        Assert.True(fi.Exists, $"Missing WC test file: {fi.FullName}");
        return IrReader.Read(new WmlDocument(fi.FullName), ReadOpts);
    }

    /// <summary>
    /// Build the (base, variant) file-name pair list from the WC directory by the documented rules.
    /// Deterministic: families and variants are sorted by ordinal name.
    /// </summary>
    public static List<(string Base, string Variant)> BuildPairs()
    {
        var files = WcDir.GetFiles("*.docx")
            .Select(f => f.Name)
            .OrderBy(n => n, StringComparer.Ordinal)
            .ToList();

        var pairs = new List<(string, string)>();
        var consumed = new HashSet<string>(StringComparer.Ordinal);

        // --- Rule 1: -Before… / -After… families.
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
                continue;

            if (befores.Count > 1)
            {
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

        // --- Rule 3: WCnnn-prefix fan-out around an -Unmodified base.
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
        fileName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) ? fileName[..^5] : fileName;

    private static void AddPair(
        List<(string, string)> pairs, HashSet<string> consumed, string baseFile, string variant)
    {
        pairs.Add((baseFile, variant));
        consumed.Add(baseFile);
        consumed.Add(variant);
    }

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
        string indexTail = stem[(idx + token.Length)..];
        return (family, isBefore, indexTail);
    }

    private static bool IsVariantOf(string baseStem, string other) =>
        Stem(other).Length > baseStem.Length &&
        Stem(other).StartsWith(baseStem + "-", StringComparison.Ordinal);

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
