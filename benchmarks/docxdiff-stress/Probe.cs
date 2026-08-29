// Sub-stage probe for IrReader: replicates the package-open / Unid / registry steps so the
// cost inside a single Read can be attributed without a native profiler.

using System.Diagnostics;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Ir;
using Docxodus.Ir.Diff;

namespace Docxodus.Stress;

internal static class Probe
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public static void Run(byte[] bytes, int n)
    {
        var doc = new WmlDocument("probe.docx", bytes);

        Console.WriteLine();
        Console.WriteLine("=== IrReader sub-stage probe (median ms of {0}) ===", n);

        Time("WmlDocument copy", n, () => { var w = new WmlDocument(doc); return w.DocumentByteArray.Length; });

        Time("package open (unzip + WordprocessingDocument)", n, () =>
        {
            using var s = new OpenXmlMemoryStreamDocument(new WmlDocument(doc));
            using var w = s.GetWordprocessingDocument();
            return w.MainDocumentPart!.Uri.ToString().Length;
        });

        Time("package open + main GetXDocument", n, () =>
        {
            using var s = new OpenXmlMemoryStreamDocument(new WmlDocument(doc));
            using var w = s.GetWordprocessingDocument();
            return w.MainDocumentPart!.GetXDocument().Root!.Name.LocalName.Length;
        });

        Time("package open + ALL reader parts GetXDocument (~HasRevisionMarkup)", n, () =>
        {
            using var s = new OpenXmlMemoryStreamDocument(new WmlDocument(doc));
            using var w = s.GetWordprocessingDocument();
            var main = w.MainDocumentPart!;
            var total = 0;
            foreach (var part in ScannableParts(main))
            {
                var root = part.GetXDocument().Root;
                if (root is null) continue;
                foreach (var e in root.DescendantsAndSelf()) total++;
            }

            return total;
        });

        // Unid assignment on a freshly parsed main part (the cold case Read always hits).
        Time("UnidHelper.AssignToAllElementsDeterministic (main)", n, () =>
        {
            using var s = new OpenXmlMemoryStreamDocument(new WmlDocument(doc));
            using var w = s.GetWordprocessingDocument();
            var root = w.MainDocumentPart!.GetXDocument().Root!;
            var t = Stopwatch.GetTimestamp();
            UnidHelper.AssignToAllElementsDeterministic(root);
            return (int)(Stopwatch.GetTimestamp() - t);
        });

        Time("IrReader.Read (RetainSources=false)", n,
            () => IrReader.Read(doc, new IrReaderOptions { RetainSources = false, RevisionView = RevisionView.Accept }).Body.Blocks.Count);

        Time("IrReader.Read (RetainSources=true)", n,
            () => IrReader.Read(doc, new IrReaderOptions { RetainSources = true, RevisionView = RevisionView.Accept }).Body.Blocks.Count);

        Time("IrReader.Read (RevisionView=FailIfPresent)", n,
            () => IrReader.Read(doc, new IrReaderOptions { RetainSources = false, RevisionView = RevisionView.FailIfPresent }).Body.Blocks.Count);

        Time("IrReader.Read (body scope only)", n,
            () => IrReader.Read(doc, new IrReaderOptions { RetainSources = false, RevisionView = RevisionView.Accept, Scopes = IrScopes.Body }).Body.Blocks.Count);

        // The floor: what any pipeline pays just to get the bytes in and a package back out.
        Time("open + reserialize package unchanged (redline floor)", n, () =>
        {
            using var sd = new OpenXmlMemoryStreamDocument(new WmlDocument(doc));
            using (var wd = sd.GetWordprocessingDocument())
            {
                _ = wd.MainDocumentPart!.GetXDocument();
            }

            return sd.GetModifiedWmlDocument().DocumentByteArray.Length;
        });

        Time("parse every reader part, no IR built (read floor)", n, () =>
        {
            using var sd = new OpenXmlMemoryStreamDocument(new WmlDocument(doc));
            using var wd = sd.GetWordprocessingDocument();
            var n2 = 0;
            foreach (var part in ScannableParts(wd.MainDocumentPart!))
                n2 += part.GetXDocument().Root?.DescendantsAndSelf().Count() ?? 0;
            return n2;
        });

        // Element counts, for context on what the walks are traversing.
        using var s2 = new OpenXmlMemoryStreamDocument(new WmlDocument(doc));
        using var w2 = s2.GetWordprocessingDocument();
        var m = w2.MainDocumentPart!;
        Console.WriteLine();
        foreach (var part in ScannableParts(m))
        {
            var root = part.GetXDocument().Root;
            if (root is null) continue;
            Console.WriteLine($"  part {part.Uri,-32} {root.DescendantsAndSelf().Count(),8:N0} elements");
        }
    }

    private static IEnumerable<OpenXmlPart> ScannableParts(MainDocumentPart main)
    {
        yield return main;
        foreach (var p in main.HeaderParts) yield return p;
        foreach (var p in main.FooterParts) yield return p;
        if (main.FootnotesPart != null) yield return main.FootnotesPart;
        if (main.EndnotesPart != null) yield return main.EndnotesPart;
        if (main.WordprocessingCommentsPart != null) yield return main.WordprocessingCommentsPart;
    }

    // Tiered JIT promotes a method only after it has been called many times, so two warm-up passes
    // leave the first cases in a series paying for JIT that the later ones inherit warm — enough to
    // report a stage as several times its steady-state cost. Every case gets the same promotion
    // budget before any of them is timed.
    private const int JitWarmup = 30;

    private static void Time<T>(string label, int n, Func<T> act)
    {
        for (var i = 0; i < JitWarmup; i++) _ = act();
        var samples = new double[n];
        for (var i = 0; i < n; i++)
        {
            var sw = Stopwatch.StartNew();
            _ = act();
            sw.Stop();
            samples[i] = sw.Elapsed.TotalMilliseconds;
        }

        Array.Sort(samples);
        Console.WriteLine($"  {label,-58} {samples[n / 2],8:F1} ms");
    }
}

internal static class UnidProbe
{
    private const int JitWarmup = 30;

    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XName Unid = "{http://powertools.codeplex.com/2011}Unid";

    public static void Run(byte[] bytes, int n)
    {
        var doc = new WmlDocument("probe.docx", bytes);

        XElement FreshRoot()
        {
            var s = new OpenXmlMemoryStreamDocument(new WmlDocument(doc));
            var w = s.GetWordprocessingDocument();
            return w.MainDocumentPart!.GetXDocument().Root!;
        }

        var root = FreshRoot();
        var all = root.DescendantsAndSelf().ToList();
        Console.WriteLine();
        Console.WriteLine($"=== Unid walk attribution ({all.Count:N0} elements, median of {n}) ===");

        Median("DescendantsAndSelf() enumeration only", n, () =>
        {
            var c = 0;
            foreach (var e in root.DescendantsAndSelf()) c++;
            return c;
        });

        Median("live-set construction (cold: every element missing)", n, () =>
        {
            var fresh = FreshRoot();
            var live = new HashSet<XElement>();
            foreach (var el in fresh.DescendantsAndSelf())
            {
                if (el.Attribute(Unid) is not null || el == fresh) continue;
                for (XElement? a = el.Parent; a is not null; a = a.Parent)
                    if (!live.Add(a)) break;
            }

            return live.Count;
        });

        Median("hasBlockDescendants scan for every element", n, () =>
        {
            var hits = 0;
            foreach (var e in all)
                if (e.Descendants().Any(d => d.Name == W + "p" || d.Name == W + "tbl")) hits++;
            return hits;
        });

        Median("w:t text concat for every element", n, () =>
        {
            var len = 0;
            foreach (var e in all)
                len += string.Concat(e.Descendants(W + "t").Select(t => (string)t)).Length;
            return len;
        });

        Median("descendant-name StringBuilder for every element", n, () =>
        {
            var len = 0;
            foreach (var e in all)
            {
                var sb = new StringBuilder(64);
                foreach (var d in e.Descendants())
                {
                    if (d.Name == W + "t") continue;
                    sb.Append(d.Name.LocalName).Append(',');
                }

                len += sb.Length;
            }

            return len;
        });

        Median($"{all.Count * 2:N0} ShortHash calls (sig + derive per element)", n, () =>
        {
            var len = 0;
            foreach (var e in all)
            {
                len += UnidHelper.ShortHash(e.Name.LocalName + "|abc|def|", 16).Length;
                len += UnidHelper.ShortHash("seed:tag:0123456789abcdef:0", 32).Length;
            }

            return len;
        });

        Median("AssignToAllElementsDeterministic (cold, full)", n, () =>
        {
            var fresh = FreshRoot();
            UnidHelper.AssignToAllElementsDeterministic(fresh);
            return 1;
        });
    }

    private static void Median<T>(string label, int n, Func<T> act)
    {
        for (var i = 0; i < JitWarmup; i++) _ = act();
        var s = new double[n];
        for (var i = 0; i < n; i++)
        {
            var sw = Stopwatch.StartNew();
            _ = act();
            sw.Stop();
            s[i] = sw.Elapsed.TotalMilliseconds;
        }

        Array.Sort(s);
        Console.WriteLine($"  {label,-58} {s[n / 2],8:F1} ms");
    }
}

// End-to-end decomposition of DocxDiff.Compare, replicating DocxDiffComparison's own steps.
internal static class PipelineProbe
{
    private const int JitWarmup = 30;

    public static void Run(byte[] leftBytes, byte[] rightBytes, int n)
    {
        var left = new WmlDocument("l.docx", leftBytes);
        var right = new WmlDocument("r.docx", rightBytes);
        var settings = new DocxDiffSettings();
        var diff = settings.ToIrDiffSettings();
        var dataDiff = diff with { CrossParagraphTokenDiff = false };
        var readOpts = DocxDiff.ReadOpts;
        var srcOpts = new IrReaderOptions { RetainSources = true, RevisionView = RevisionView.Accept };

        Console.WriteLine();
        Console.WriteLine($"=== DocxDiff.Compare decomposition (median of {n}) ===");

        Median("StrictOoxmlNormalizer.NormalizeToTransitional x2", n, () =>
        {
            _ = StrictOoxmlNormalizer.NormalizeToTransitional(left);
            return StrictOoxmlNormalizer.NormalizeToTransitional(right);
        });

        Median("MarkupCompatibilityNormalizer.Normalize x2", n, () =>
        {
            _ = MarkupCompatibilityNormalizer.Normalize(left);
            return MarkupCompatibilityNormalizer.Normalize(right);
        });

        Median("DocxDiff.PreAccept x2 (both normalizers)", n, () =>
        {
            _ = DocxDiff.PreAccept(settings, left);
            return DocxDiff.PreAccept(settings, right);
        });

        var preLeft = DocxDiff.PreAccept(settings, left);
        var preRight = DocxDiff.PreAccept(settings, right);

        Median("IrReader.Read x2 (script opts, sequential)", n, () =>
        {
            _ = IrReader.Read(preLeft, readOpts);
            return IrReader.Read(preRight, readOpts);
        });

        Median("IrReader.Read x2 (script opts, PARALLEL)", n, () =>
        {
            IrDocument? a = null, b = null;
            var t = Task.Run(() => a = IrReader.Read(preLeft, readOpts));
            b = IrReader.Read(preRight, readOpts);
            t.Wait();
            return (a, b);
        });

        Median("IrReader.Read x2 (RetainSources, sequential)", n, () =>
        {
            _ = IrReader.Read(preLeft, srcOpts);
            return IrReader.Read(preRight, srcOpts);
        });

        var irLeft = IrReader.Read(preLeft, readOpts);
        var irRight = IrReader.Read(preRight, readOpts);

        Median("IrEditScriptBuilder.Build", n, () => IrEditScriptBuilder.Build(irLeft, irRight, dataDiff));

        var script = IrEditScriptBuilder.Build(irLeft, irRight, dataDiff);

        Median("IrMarkupRenderer.Render (incl. its 2 internal reads)", n,
            () => IrMarkupRenderer.Render(script, preLeft, preRight, diff));

        Median("DocxDiff.Compare (whole thing)", n, () => DocxDiff.Compare(left, right));
    }

    private static void Median<T>(string label, int n, Func<T> act)
    {
        for (var i = 0; i < JitWarmup; i++) _ = act();
        var s = new double[n];
        for (var i = 0; i < n; i++)
        {
            var sw = Stopwatch.StartNew();
            _ = act();
            sw.Stop();
            s[i] = sw.Elapsed.TotalMilliseconds;
        }

        Array.Sort(s);
        Console.WriteLine($"  {label,-58} {s[n / 2],8:F1} ms");
    }
}

// How repetitive are the strings UnidHelper hashes? Decides whether memoizing ShortHash pays.
internal static class SigProbe
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public static void Run(byte[] bytes)
    {
        var doc = new WmlDocument("probe.docx", bytes);
        using var s = new OpenXmlMemoryStreamDocument(new WmlDocument(doc));
        using var w = s.GetWordprocessingDocument();

        Console.WriteLine();
        Console.WriteLine("=== ContentSignature input repetition ===");
        foreach (var (name, part) in Parts(w))
        {
            var root = part?.GetXDocument().Root;
            if (root is null) continue;
            var sigs = new List<string>();
            foreach (var e in root.DescendantsAndSelf()) sigs.Add(SigInput(e));
            var distinct = sigs.Distinct(StringComparer.Ordinal).Count();
            Console.WriteLine($"  {name,-14} {sigs.Count,7:N0} elements   {distinct,7:N0} distinct signature inputs   {100.0 * (sigs.Count - distinct) / Math.Max(1, sigs.Count),5:F1}% repeats");
        }
    }

    private static IEnumerable<(string, OpenXmlPart?)> Parts(WordprocessingDocument w)
    {
        yield return ("document", w.MainDocumentPart);
        yield return ("footnotes", w.MainDocumentPart?.FootnotesPart);
    }

    // Mirrors UnidHelper.ContentSignature's hashed input (not its output).
    private static string SigInput(XElement element)
    {
        if (element.Descendants().Any(d => d.Name == W + "p" || d.Name == W + "tbl"))
            return element.Name.LocalName;

        var text = string.Concat(element.Descendants(W + "t").Select(t => (string)t));
        var pPr = element.Element(W + "pPr");
        var styleId = pPr?.Element(W + "pStyle")?.Attribute(W + "val")?.Value ?? string.Empty;
        var numId = pPr?.Element(W + "numPr")?.Element(W + "numId")?.Attribute(W + "val")?.Value ?? string.Empty;
        var sb = new StringBuilder(text.Length + 64);
        sb.Append(text).Append('|').Append(styleId).Append('|').Append(numId).Append('|');
        foreach (var d in element.Descendants())
        {
            if (d.Name == W + "t") continue;
            sb.Append(d.Name.LocalName).Append(',');
        }

        return sb.ToString();
    }
}
