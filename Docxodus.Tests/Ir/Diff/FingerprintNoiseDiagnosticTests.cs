#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;
using Xunit.Abstractions;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// THROWAWAY diagnostic (M2.2 Task 4, sub-task B step 1) over the WC-BodyBookmarks pair — the sole
/// source of the corpus' 1,714 FormatOnly entries (M2.1 carry-list). It reads both sides with
/// <c>RetainSources=true</c>, pairs ContentHash-equal / FormatFingerprint-different paragraphs, and for
/// a handful dumps (a) the modeled per-run field differences and (b) the leftover (unmodeled) rPr XML
/// that survives into <see cref="IrRunFormat.UnmodeledDigest"/>. The dump identifies the actual noise
/// elements so the resolution (ModeledOnly format comparison vs. an IR normalization rule) is evidence-
/// based. Kept as a <c>Diagnostic</c>-trait test (not run by default) for reproducibility.
/// </summary>
[Trait("Category", "Diagnostic")]
public class FingerprintNoiseDiagnosticTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // Mirror IrReader.RPrConsumed (the always-modeled rPr children). Everything else in an rPr lands in
    // the unmodeled digest. w:rFonts is only partially consumed (ascii) so it is NOT in this set.
    private static readonly HashSet<XName> Consumed = new()
    {
        W + "rStyle", W + "b", W + "i", W + "strike", W + "dstrike", W + "caps",
        W + "smallCaps", W + "vanish", W + "u", W + "sz", W + "color", W + "highlight", W + "shd",
    };

    private static readonly IrReaderOptions ReadOptsRetain =
        new() { RetainSources = true, RevisionView = RevisionView.Accept };

    private readonly ITestOutputHelper _out;

    public FingerprintNoiseDiagnosticTests(ITestOutputHelper output) => _out = output;

    [Fact]
    public void Dump_BodyBookmarks_format_noise_inventory()
    {
        var left = ReadRetain("WC-BodyBookmarks-Before.docx");
        var right = ReadRetain("WC-BodyBookmarks-After.docx");

        // Pair content-equal / format-different paragraphs positionally over the FormatOnly aligner entries.
        var alignment = IrBlockAligner.Align(left, right, new IrDiffSettings());
        var formatOnlyPairs = alignment.Entries
            .Where(e => e.Kind == IrAlignmentKind.FormatOnly
                        && e.Left is IrParagraph && e.Right is IrParagraph)
            .Select(e => ((IrParagraph)e.Left!, (IrParagraph)e.Right!))
            .ToList();

        _out.WriteLine($"FormatOnly paragraph pairs in WC-BodyBookmarks: {formatOnlyPairs.Count}");
        _out.WriteLine("");

        // Tally which leftover (unmodeled) rPr child names appear in any FormatOnly run, across ALL pairs.
        var leftoverNameTally = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var (lp, rp) in formatOnlyPairs)
        {
            foreach (var rPr in LeftoverRPrs(lp).Concat(LeftoverRPrs(rp)))
                foreach (var name in rPr.Elements().Select(e => e.Name.LocalName))
                    leftoverNameTally[name] = leftoverNameTally.TryGetValue(name, out int c) ? c + 1 : 1;
        }

        _out.WriteLine("Leftover (unmodeled) rPr child elements across ALL FormatOnly pairs (name → occurrences):");
        foreach (var kv in leftoverNameTally.OrderByDescending(k => k.Value).ThenBy(k => k.Key, StringComparer.Ordinal))
            _out.WriteLine($"  w:{kv.Key} = {kv.Value}");
        _out.WriteLine("");

        // Detailed dump of the first 5 pairs: modeled field diffs + leftover rPr XML per run.
        foreach (var (lp, rp) in formatOnlyPairs.Take(5))
        {
            _out.WriteLine($"=== pair {lp.Anchor} ↔ {rp.Anchor} ===");
            var lRuns = Runs(lp).ToList();
            var rRuns = Runs(rp).ToList();
            _out.WriteLine($"  left runs={lRuns.Count} right runs={rRuns.Count}");

            int n = Math.Max(lRuns.Count, rRuns.Count);
            for (int i = 0; i < n; i++)
            {
                var lf = i < lRuns.Count ? lRuns[i].Format : null;
                var rf = i < rRuns.Count ? rRuns[i].Format : null;
                var modeledDiff = ModeledFieldDiff(lf, rf);
                if (modeledDiff.Count > 0)
                    _out.WriteLine($"  run[{i}] modeled-diff: {string.Join(", ", modeledDiff)}");
            }

            DumpLeftovers("  LEFT leftover rPr", lp);
            DumpLeftovers("  RIGHT leftover rPr", rp);
            _out.WriteLine("");
        }
    }

    private void DumpLeftovers(string label, IrParagraph p)
    {
        var leftovers = LeftoverRPrs(p)
            .Select(rPr => rPr.ToString(SaveOptions.DisableFormatting))
            .Where(s => s.Length > 0)
            .ToList();
        if (leftovers.Count == 0)
            return;
        _out.WriteLine($"{label} ({leftovers.Count} runs with leftovers):");
        foreach (var xml in leftovers.Take(6))
            _out.WriteLine($"    {xml}");
    }

    /// <summary>
    /// The leftover-only <c>w:rPr</c> elements for every text-bearing run in the paragraph: each source
    /// run's rPr cloned with the modeled (RPrConsumed) children removed, so what remains is exactly what
    /// the <see cref="IrRunFormat.UnmodeledDigest"/> hashed. Empty leftover rPrs are dropped.
    /// </summary>
    private static IEnumerable<XElement> LeftoverRPrs(IrParagraph p)
    {
        var pElement = p.Source.Element;
        if (pElement is null)
            yield break;
        foreach (var r in pElement.Elements(W + "r"))
        {
            var rPr = r.Element(W + "rPr");
            if (rPr is null)
                continue;
            var clone = new XElement(rPr);
            clone.Elements().Where(e => Consumed.Contains(e.Name)).Remove();
            // w:vertAlign is consumed only for sub/superscript; a baseline vertAlign stays unmodeled — leave it.
            if (clone.HasElements)
                yield return clone;
        }
    }

    private static IEnumerable<IrTextRun> Runs(IrParagraph p) =>
        p.Inlines.OfType<IrTextRun>();

    private static List<string> ModeledFieldDiff(IrRunFormat? a, IrRunFormat? b)
    {
        var diffs = new List<string>();
        void Cmp<T>(string name, T? av, T? bv)
        {
            if (!EqualityComparer<T?>.Default.Equals(av, bv))
                diffs.Add($"{name}({av?.ToString() ?? "∅"}→{bv?.ToString() ?? "∅"})");
        }
        if (a is null || b is null)
        {
            diffs.Add(a is null ? "<left run absent>" : "<right run absent>");
            return diffs;
        }
        Cmp("StyleId", a.StyleId, b.StyleId);
        Cmp("Bold", a.Bold, b.Bold);
        Cmp("Italic", a.Italic, b.Italic);
        Cmp("Strike", a.Strike, b.Strike);
        Cmp("FontAscii", a.FontAscii, b.FontAscii);
        Cmp("SizeHalfPoints", a.SizeHalfPoints, b.SizeHalfPoints);
        Cmp("ColorHex", a.ColorHex, b.ColorHex);
        Cmp("Highlight", a.Highlight, b.Highlight);
        Cmp("Caps", a.Caps, b.Caps);
        Cmp("SmallCaps", a.SmallCaps, b.SmallCaps);
        Cmp("Vanish", a.Vanish, b.Vanish);
        Cmp("UnmodeledDigest", a.UnmodeledDigest.ToHex(), b.UnmodeledDigest.ToHex());
        return diffs;
    }

    private static IrDocument ReadRetain(string fileName)
    {
        var fi = new System.IO.FileInfo(System.IO.Path.Combine(WcCorpus.WcDir.FullName, fileName));
        Assert.True(fi.Exists, $"Missing WC test file: {fi.FullName}");
        return IrReader.Read(new Docxodus.WmlDocument(fi.FullName), ReadOptsRetain);
    }
}
