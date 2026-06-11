#nullable enable

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Ir;
using Xunit;
using Xunit.Abstractions;

namespace Docxodus.Tests.Ir;

/// <summary>
/// Phase 1 gate criterion 2 (perf budget). Compares corpus-wide wall time for the shipped
/// <see cref="WmlToMarkdownConverter.Convert"/> (the oracle) against the IR path
/// (<see cref="IrReader.Read"/> + <see cref="IrMarkdownEmitter.Emit"/>) over the SAME prepared
/// inputs the equivalence harness uses, and asserts the IR path stays within a tolerant 2.0× of the
/// oracle. Both totals are printed via <see cref="ITestOutputHelper"/>.
///
/// <para>
/// The memory measurement is REPORTED, not asserted: a working-set / allocation delta is too flaky
/// for CI (GC timing, shared corpus state). We measure managed allocated bytes for one IR snapshot of
/// the largest fixture (<see cref="GC.GetTotalAllocatedBytes"/> delta) and report the ratio against
/// the document's main-part XML size. See the gate report for methodology + the recorded number.
/// </para>
/// <para>
/// The full corpus benchmark is OPT-IN (<c>DOCXODUS_RUN_PERF=1</c>): it forces blocking full GCs and
/// churns hundreds of MB which, run concurrently in the default parallel suite, starves the SkiaSharp
/// native image-rendering tests and flakes them. Opt-in also gives an uncontended measurement. The
/// default run executes a fast, GC-quiet <see cref="SmokeCheck"/> (order-of-magnitude guard only).
/// </para>
/// </summary>
public class IrMarkdownPerfBudgetTests
{
    private static readonly DirectoryInfo TestFilesDir = new("../../../../TestFiles/");

    // Tolerant: warm-up pass first, Stopwatch totals, generous bound (CI jitter, cold JIT, shared box).
    private const double MaxIrToOracleRatio = 2.0;

    private readonly ITestOutputHelper _output;

    public IrMarkdownPerfBudgetTests(ITestOutputHelper output) => _output = output;

    // The full corpus benchmark forces blocking full GCs and churns hundreds of MB; run concurrently
    // (xUnit parallelizes across test classes by default) it starves the SkiaSharp native image-
    // rendering tests (OxPt.HcTests image/RTL fixtures) and flakes them. CI runs the whole suite
    // unfiltered, so the heavy path is OPT-IN via DOCXODUS_RUN_PERF=1 (the gate is run explicitly,
    // also giving an uncontended measurement). The default run executes a fast, GC-quiet
    // handful-of-fixtures smoke check that still asserts the IR path is in the same order of magnitude
    // as the oracle, so the gate never silently rots.
    private static bool RunFullBenchmark =>
        string.Equals(Environment.GetEnvironmentVariable("DOCXODUS_RUN_PERF"), "1", StringComparison.Ordinal);

    [Fact]
    [Trait("Category", "Perf")]
    public void IrPath_WallTime_WithinBudget_OfOracle()
    {
        if (!RunFullBenchmark)
        {
            SmokeCheck();
            return;
        }

        var files = TestFilesDir.GetFiles("*.docx", SearchOption.AllDirectories)
            .OrderBy(f => f.FullName, StringComparer.Ordinal)
            .ToList();

        // Prepare ALL inputs up front (revision-accept once) so neither path pays for I/O or
        // revision acceptance inside the timed region — we are measuring read+emit vs convert only.
        var prepared = new List<byte[]>();
        foreach (var file in files)
        {
            if (!CanOpen(file)) continue;
            try { prepared.Add(PrepareInput(file).DocumentByteArray); }
            catch { /* skip fixtures the prepare step rejects, exactly like the equivalence harness */ }
        }

        Assert.NotEmpty(prepared);

        // Warm-up: one full pass over both paths so JIT/first-touch costs do not skew the totals.
        RunOraclePass(prepared);
        RunIrPass(prepared);

        // Take the best (minimum) of a few full passes for each path. The minimum is the least
        // noise-contaminated estimate of the true cost (jitter only ever adds time), so best-of-N
        // de-flakes the ratio on a shared CI box without inflating either side artificially.
        const int passes = 3;
        double oracleMs = double.MaxValue, irMs = double.MaxValue;
        for (int i = 0; i < passes; i++)
        {
            var sw = Stopwatch.StartNew();
            RunOraclePass(prepared);
            sw.Stop();
            oracleMs = Math.Min(oracleMs, sw.Elapsed.TotalMilliseconds);
        }
        for (int i = 0; i < passes; i++)
        {
            var sw = Stopwatch.StartNew();
            RunIrPass(prepared);
            sw.Stop();
            irMs = Math.Min(irMs, sw.Elapsed.TotalMilliseconds);
        }

        var ratio = irMs / oracleMs;

        _output.WriteLine($"Corpus fixtures timed: {prepared.Count}");
        _output.WriteLine($"Oracle (WmlToMarkdownConverter.Convert) total: {oracleMs:F1} ms");
        _output.WriteLine($"IR path (IrReader.Read + IrMarkdownEmitter.Emit) total: {irMs:F1} ms");
        _output.WriteLine($"IR / oracle ratio: {ratio:F2}× (budget ≤ {MaxIrToOracleRatio:F1}×)");

        // Memory spot-check — REPORTED only (see class doc): managed allocation for one IR snapshot of
        // the largest fixture vs that document's main-part XML size.
        ReportMemorySpotCheck();

        Assert.True(ratio <= MaxIrToOracleRatio,
            $"IR path wall time {irMs:F1} ms is {ratio:F2}× the oracle's {oracleMs:F1} ms, " +
            $"exceeding the {MaxIrToOracleRatio:F1}× budget.");
    }

    /// <summary>
    /// Fast, GC-quiet default-run check: time the oracle and IR paths over a handful of small
    /// fixtures (no corpus loop, no forced full GCs — so it cannot starve concurrent SkiaSharp tests)
    /// and assert the IR path stays within a lenient bound. Keeps the gate from silently rotting
    /// without the heavy benchmark's parallel-run hazards. The authoritative measurement is the
    /// opt-in full benchmark (DOCXODUS_RUN_PERF=1); the recorded numbers live in the gate report.
    /// </summary>
    private void SmokeCheck()
    {
        var smallFixtures = new[]
        {
            "CA001-Plain.docx", "CZ002-Multi-Paragraphs.docx", "CA005-Table.docx",
            "CA003-Numbered-List.docx", "HC039-Bold.docx",
        };

        var prepared = new List<byte[]>();
        foreach (var name in smallFixtures)
        {
            var hit = TestFilesDir.GetFiles(name, SearchOption.AllDirectories).FirstOrDefault();
            if (hit is not null && CanOpen(hit)) prepared.Add(PrepareInput(hit).DocumentByteArray);
        }
        Assert.NotEmpty(prepared);

        // Warm-up, then best-of-3 (minimum) for each path to shed jitter without forcing GC.
        RunOraclePass(prepared);
        RunIrPass(prepared);
        double oracleMs = double.MaxValue, irMs = double.MaxValue;
        for (int i = 0; i < 3; i++)
        {
            var sw = Stopwatch.StartNew(); RunOraclePass(prepared); sw.Stop();
            oracleMs = Math.Min(oracleMs, sw.Elapsed.TotalMilliseconds);
        }
        for (int i = 0; i < 3; i++)
        {
            var sw = Stopwatch.StartNew(); RunIrPass(prepared); sw.Stop();
            irMs = Math.Min(irMs, sw.Elapsed.TotalMilliseconds);
        }

        var ratio = irMs / oracleMs;
        _output.WriteLine($"[smoke] oracle {oracleMs:F1} ms / IR {irMs:F1} ms = {ratio:F2}× " +
                          $"({prepared.Count} small fixtures). Full benchmark: set DOCXODUS_RUN_PERF=1.");

        // Lenient: the few-fixture smoke ratio is noisier than the corpus ratio AND inflated by the
        // IR's fixed per-open overhead (registry construction) dominating on tiny bodies — it routinely
        // sits ~3-4×. The point here is only to catch an ORDER-OF-MAGNITUDE regression, not to
        // re-assert the 2.0× corpus budget (that is the opt-in full benchmark's job), so the bound is
        // deliberately loose.
        const double smokeBound = 8.0;
        Assert.True(ratio <= smokeBound,
            $"[smoke] IR path {irMs:F1} ms is {ratio:F2}× the oracle's {oracleMs:F1} ms (> {smokeBound:F1}× " +
            "smoke bound). Run the full benchmark (DOCXODUS_RUN_PERF=1) to confirm against the 2.0× corpus budget.");
    }

    private static void RunOraclePass(List<byte[]> prepared)
    {
        foreach (var bytes in prepared)
        {
            var doc = new WmlDocument("perf.docx", bytes);
            _ = WmlToMarkdownConverter.Convert(doc, new WmlToMarkdownConverterSettings());
        }
    }

    private static void RunIrPass(List<byte[]> prepared)
    {
        foreach (var bytes in prepared)
        {
            var doc = new WmlDocument("perf.docx", bytes);
            var ir = IrReader.Read(doc);
            _ = IrMarkdownEmitter.Emit(ir, new WmlToMarkdownConverterSettings());
        }
    }

    private void ReportMemorySpotCheck()
    {
        // The 3× reference is against the DOCUMENT XML the IR models, so pick the fixture with the
        // largest MAIN-PART XML — NOT the largest .docx file (whose bulk is usually embedded images /
        // glossary parts the IR does not snapshot, which makes the file-size proxy meaningless).
        FileInfo? largest = null;
        long largestXml = 0;
        foreach (var file in TestFilesDir.GetFiles("*.docx", SearchOption.AllDirectories).Where(CanOpen))
        {
            try
            {
                var size = MainPartXmlSize(PrepareInput(file).DocumentByteArray);
                if (size > largestXml) { largestXml = size; largest = file; }
            }
            catch { /* skip prepare-rejected fixtures */ }
        }

        Assert.NotNull(largest);
        var bytes = PrepareInput(largest!).DocumentByteArray;
        var xmlBytes = MainPartXmlSize(bytes);

        // Two complementary numbers per mode, all reported (none asserted — see class doc):
        //  - RETAINED (live-heap delta with the snapshot rooted; GC.GetTotalMemory(true) brackets) —
        //    the closest cheap proxy for the snapshot's resident footprint, which is what the X-ratio
        //    reference is about.
        //  - CHURN (GC.GetTotalAllocatedBytes delta) — total managed bytes allocated during the read;
        //    includes transient XML/parse garbage, so it over-counts the snapshot but bounds peak
        //    pressure.
        // Measured for BOTH provenance modes (M1.5 Task 2): RetainSources=true pins the parsed XML
        // (Sources + per-node IrProvenance.Element), false drops it (Sources empty, Element null) so the
        // working XDocuments become collectible after Read. PartUri facts survive in both.
        var on = MeasureSnapshot(bytes, new IrReaderOptions { RetainSources = true });
        var off = MeasureSnapshot(bytes, new IrReaderOptions { RetainSources = false });

        _output.WriteLine($"Memory spot-check (largest-body fixture: {largest!.Name}):");
        _output.WriteLine($"  main-part XML size: {xmlBytes:N0} bytes");
        _output.WriteLine($"  RetainSources=true  RETAINED (live-heap delta): {on.Retained:N0} bytes ({Ratio(on.Retained, xmlBytes):F2}× XML)");
        _output.WriteLine($"  RetainSources=true  CHURN    (alloc delta):     {on.Churn:N0} bytes ({Ratio(on.Churn, xmlBytes):F2}× XML)");
        _output.WriteLine($"  RetainSources=false RETAINED (live-heap delta): {off.Retained:N0} bytes ({Ratio(off.Retained, xmlBytes):F2}× XML)");
        _output.WriteLine($"  RetainSources=false CHURN    (alloc delta):     {off.Churn:N0} bytes ({Ratio(off.Churn, xmlBytes):F2}× XML)");
        _output.WriteLine("  (the X-ratio reference is the retained-mode live-heap ratio; reported, not asserted — see gate report methodology)");
    }

    private static double Ratio(long value, long xmlBytes) =>
        xmlBytes > 0 ? (double)value / xmlBytes : double.NaN;

    /// <summary>One IR snapshot's resident (live-heap delta, snapshot rooted) and churn (total-allocated
    /// delta) bytes, measured with the same GC bracketing the class doc describes.</summary>
    private static (long Retained, long Churn) MeasureSnapshot(byte[] bytes, IrReaderOptions options)
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        var liveBefore = GC.GetTotalMemory(forceFullCollection: true);
        var allocBefore = GC.GetTotalAllocatedBytes(precise: true);
        var ir = IrReader.Read(new WmlDocument("perf.docx", bytes), options);
        var allocAfter = GC.GetTotalAllocatedBytes(precise: true);
        var liveAfter = GC.GetTotalMemory(forceFullCollection: true);
        GC.KeepAlive(ir);
        return (liveAfter - liveBefore, allocAfter - allocBefore);
    }

    private static long MainPartXmlSize(byte[] bytes)
    {
        using var ms = new MemoryStream(bytes, writable: false);
        using var wdoc = WordprocessingDocument.Open(ms, false);
        var main = wdoc.MainDocumentPart;
        if (main is null) return 0;
        using var partStream = main.GetStream();
        return partStream.Length;
    }

    // --- prepared-input parity with the equivalence harness -------------------------------------

    private static readonly string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static readonly XName[] RevisionElementNames =
    {
        XName.Get("ins", W), XName.Get("del", W), XName.Get("moveFrom", W), XName.Get("moveTo", W),
        XName.Get("rPrChange", W), XName.Get("pPrChange", W), XName.Get("tblPrChange", W),
        XName.Get("trPrChange", W), XName.Get("tcPrChange", W), XName.Get("sectPrChange", W),
        XName.Get("numberingChange", W),
    };

    private static WmlDocument PrepareInput(FileInfo file)
    {
        var doc = new WmlDocument(file.FullName);
        return HasRevisionMarkup(doc) ? RevisionProcessor.AcceptRevisions(doc) : doc;
    }

    private static bool HasRevisionMarkup(WmlDocument doc)
    {
        try
        {
            using var stream = new OpenXmlMemoryStreamDocument(doc);
            using var wdoc = stream.GetWordprocessingDocument();
            var names = new HashSet<XName>(RevisionElementNames);
            foreach (var root in ScopeRoots(wdoc))
                if (root is not null && root.DescendantsAndSelf().Any(e => names.Contains(e.Name)))
                    return true;
            return false;
        }
        catch
        {
            return false;
        }
    }

    private static IEnumerable<XElement?> ScopeRoots(WordprocessingDocument wdoc)
    {
        var main = wdoc.MainDocumentPart;
        if (main is null) yield break;
        yield return main.GetXDocument().Root;
        foreach (var h in main.HeaderParts) yield return h.GetXDocument().Root;
        foreach (var f in main.FooterParts) yield return f.GetXDocument().Root;
        if (main.FootnotesPart is not null) yield return main.FootnotesPart.GetXDocument().Root;
        if (main.EndnotesPart is not null) yield return main.EndnotesPart.GetXDocument().Root;
    }

    private static bool CanOpen(FileInfo file)
    {
        try
        {
            using var fs = file.OpenRead();
            using var _ = WordprocessingDocument.Open(fs, false);
            return true;
        }
        catch
        {
            return false;
        }
    }
}
