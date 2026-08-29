// DocxDiff performance stress harness.
//
// Generates a family of deterministic edited variants of one heavyweight .docx and times
// the DocxDiff pipeline against each — end to end and stage by stage (pre-accept, IR read,
// edit-script build, markup render). The variants span the shapes that dominate real
// redline work and the shapes that stress the aligner hardest:
//
//   identical   no change at all (fast-path floor)
//   light       a handful of scattered word edits (a typical counsel pass)
//   heavy       an edit in every fifth paragraph
//   churn       an edit in roughly half of all text nodes
//   reorder     blocks relocated across the document (move detection)
//   structural  paragraphs deleted and new ones inserted
//   footnotes   edits inside the footnote part only
//   rewrite     every paragraph's text replaced (worst case)
//
// Each case also records SHA-256 digests of the redline bytes, the revision list, and the
// edit-script JSON, so an optimization pass can be proven output-identical: run with
// --baseline before the change, --check after.
//
// The document is not committed — pass any comparable .docx on the command line.

using System.Diagnostics;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Docxodus;
using Docxodus.Ir;
using Docxodus.Ir.Diff;

const string Usage = """
usage: DocxDiffStress <document.docx> [options]
  --iterations N     timed iterations per case (default 5)
  --warmup N         untimed warm-up iterations per case (default 2)
  --cases a,b,c      restrict to named cases (default: all)
  --stages           also report per-stage timings for each case
  --products         time GetRevisions / GetEditScriptJson / the fused pass too
  --baseline FILE    write output digests to FILE
  --check FILE       compare output digests against FILE and fail on any mismatch
  --out DIR          write the generated variants here (default: skip)
""";

if (args.Length < 1 || args[0].StartsWith("--"))
{
    Console.Error.WriteLine(Usage);
    return 1;
}

var docPath = args[0];
var iterations = IntArg("--iterations", 5);
var warmup = IntArg("--warmup", 2);
var wantStages = args.Contains("--stages");
var wantProducts = args.Contains("--products");
var baselineOut = StrArg("--baseline");
var checkAgainst = StrArg("--check");
var outDir = StrArg("--out");
var caseFilter = StrArg("--cases")?.Split(',', StringSplitOptions.RemoveEmptyEntries).ToHashSet(StringComparer.OrdinalIgnoreCase);

if (outDir != null) Directory.CreateDirectory(outDir);

var baseBytes = File.ReadAllBytes(docPath);
Console.WriteLine($"document : {docPath} ({baseBytes.Length:N0} bytes)");
Console.WriteLine($"runtime  : .NET {Environment.Version}, {Environment.ProcessorCount} cores, server GC={System.Runtime.GCSettings.IsServerGC}");
Console.WriteLine($"schedule : {warmup} warm-up + {iterations} timed iterations per case");
Console.WriteLine();

if (args.Contains("--probe"))
{
    Docxodus.Stress.Probe.Run(baseBytes, iterations);
    Docxodus.Stress.UnidProbe.Run(baseBytes, iterations);
    Docxodus.Stress.SigProbe.Run(baseBytes);
    var probeRight = Docxodus.Stress.Variants.Build(baseBytes).First(v => v.Name == "light").Bytes;
    Docxodus.Stress.PipelineProbe.Run(baseBytes, probeRight, iterations);
    return 0;
}

var variants = Docxodus.Stress.Variants.Build(baseBytes);
if (caseFilter != null) variants = variants.Where(v => caseFilter.Contains(v.Name)).ToList();
if (variants.Count == 0) { Console.Error.WriteLine("no cases selected"); return 1; }

foreach (var v in variants)
{
    Console.WriteLine($"case {v.Name,-11} {v.Description} ({v.Bytes.Length:N0} bytes)");
    if (outDir != null) File.WriteAllBytes(Path.Combine(outDir, $"{v.Name}.docx"), v.Bytes);
}
Console.WriteLine();

var digests = new SortedDictionary<string, string>(StringComparer.Ordinal);
var rows = new List<Row>();

foreach (var v in variants)
{
    var left = new WmlDocument("baseline.docx", baseBytes);
    var right = new WmlDocument($"{v.Name}.docx", v.Bytes);

    for (var i = 0; i < warmup; i++) _ = DocxDiff.Compare(left, right);

    var row = new Row(v.Name, Measure(iterations, () => DocxDiff.Compare(left, right)));

    if (wantProducts)
    {
        for (var i = 0; i < warmup; i++) { _ = DocxDiff.GetRevisions(left, right); _ = DocxDiff.GetEditScriptJson(left, right); }
        row.Revisions = Measure(iterations, () => DocxDiff.GetRevisions(left, right));
        row.EditScript = Measure(iterations, () => DocxDiff.GetEditScriptJson(left, right));
        row.AllProducts = Measure(iterations, () =>
        {
            var c = DocxDiff.CreateComparison(left, right);
            _ = c.ToRedline(); _ = c.GetRevisions(); _ = c.GetEditScriptJson();
            return c;
        });
    }

    if (wantStages) row.Stages = MeasureStages(left, right, iterations, warmup);

    rows.Add(row);

    var redline = DocxDiff.Compare(left, right);
    var revs = DocxDiff.GetRevisions(left, right);
    digests[$"{v.Name}/redline"] = Sha(redline.DocumentByteArray);
    digests[$"{v.Name}/revisions"] = Sha(Encoding.UTF8.GetBytes(RevisionsFingerprint(revs)));
    digests[$"{v.Name}/editscript"] = Sha(Encoding.UTF8.GetBytes(DocxDiff.GetEditScriptJson(left, right)));
    digests[$"{v.Name}/revcount"] = revs.Count.ToString(CultureInfo.InvariantCulture);
}

Report(rows, wantProducts, wantStages);

if (baselineOut != null)
{
    File.WriteAllText(baselineOut, JsonSerializer.Serialize(digests, new JsonSerializerOptions { WriteIndented = true }));
    Console.WriteLine($"{Environment.NewLine}baseline digests written to {baselineOut} ({digests.Count} entries)");
}

var exit = 0;
if (checkAgainst != null)
{
    var expected = JsonSerializer.Deserialize<SortedDictionary<string, string>>(File.ReadAllText(checkAgainst))!;
    var mismatches = 0;
    foreach (var (k, want) in expected)
    {
        if (!digests.TryGetValue(k, out var got)) { Console.WriteLine($"[parity] MISSING  {k}"); mismatches++; }
        else if (got != want) { Console.WriteLine($"[parity] MISMATCH {k}"); Console.WriteLine($"           expected {want}"); Console.WriteLine($"           actual   {got}"); mismatches++; }
    }
    foreach (var k in digests.Keys.Where(k => !expected.ContainsKey(k))) Console.WriteLine($"[parity] EXTRA    {k}");
    Console.WriteLine();
    Console.WriteLine(mismatches == 0
        ? $"[parity] OK - all {expected.Count} digests identical to baseline"
        : $"[parity] {mismatches} MISMATCH(ES) - output changed");
    exit = mismatches == 0 ? 0 : 2;
}

return exit;

// ---------------- measurement ----------------

static Stat Measure<T>(int n, Func<T> act)
{
    var samples = new double[n];
    var alloc0 = GC.GetTotalAllocatedBytes(precise: false);
    for (var i = 0; i < n; i++)
    {
        var sw = Stopwatch.StartNew();
        _ = act();
        sw.Stop();
        samples[i] = sw.Elapsed.TotalMilliseconds;
    }
    var alloc = (GC.GetTotalAllocatedBytes(precise: false) - alloc0) / (double)n;
    Array.Sort(samples);
    return new Stat(samples[0], samples[n / 2], samples[^1], samples.Average(), alloc);
}

static StageTimes MeasureStages(WmlDocument left, WmlDocument right, int n, int warmup)
{
    var settings = new DocxDiffSettings();
    var diff = settings.ToIrDiffSettings();
    var dataDiff = diff with { CrossParagraphTokenDiff = false };
    var readOpts = DocxDiff.ReadOpts;
    var srcOpts = new IrReaderOptions { RetainSources = true, RevisionView = RevisionView.Accept };

    // Warm every stage before timing any of it.
    for (var i = 0; i < warmup; i++)
    {
        var wl = IrReader.Read(left, readOpts);
        var wr = IrReader.Read(right, readOpts);
        var ws = IrEditScriptBuilder.Build(wl, wr, dataDiff);
        _ = IrMarkupRenderer.Render(ws, left, right, diff);
    }

    var readLeft = Measure(n, () => IrReader.Read(left, readOpts));
    var readRight = Measure(n, () => IrReader.Read(right, readOpts));
    var readSrcLeft = Measure(n, () => IrReader.Read(left, srcOpts));
    var readSrcRight = Measure(n, () => IrReader.Read(right, srcOpts));

    var irLeft = IrReader.Read(left, readOpts);
    var irRight = IrReader.Read(right, readOpts);
    var build = Measure(n, () => IrEditScriptBuilder.Build(irLeft, irRight, dataDiff));

    var script = IrEditScriptBuilder.Build(irLeft, irRight, dataDiff);
    var render = Measure(n, () => IrMarkupRenderer.Render(script, left, right, diff));
    var revRender = Measure(n, () => IrRevisionRenderer.Render(script, irLeft, irRight, dataDiff));

    return new StageTimes(readLeft, readRight, readSrcLeft, readSrcRight, build, render, revRender);
}

static void Report(List<Row> rows, bool products, bool stages)
{
    Console.WriteLine();
    Console.WriteLine("=== DocxDiff.Compare (end to end) ===");
    Console.WriteLine($"{"case",-12} {"min ms",9} {"median",9} {"max ms",9} {"diffs/s",9} {"alloc MB",10}");
    foreach (var r in rows)
        Console.WriteLine($"{r.Name,-12} {r.Compare.Min,9:F1} {r.Compare.Median,9:F1} {r.Compare.Max,9:F1} {1000 / r.Compare.Median,9:F2} {r.Compare.AllocBytes / 1048576.0,10:F1}");

    if (products)
    {
        Console.WriteLine();
        Console.WriteLine("=== products (median ms) ===");
        Console.WriteLine($"{"case",-12} {"compare",9} {"revisions",10} {"editscript",11} {"fused x3",11}");
        foreach (var r in rows)
            Console.WriteLine($"{r.Name,-12} {r.Compare.Median,9:F1} {r.Revisions!.Median,10:F1} {r.EditScript!.Median,11:F1} {r.AllProducts!.Median,11:F1}");
    }

    if (stages)
    {
        Console.WriteLine();
        Console.WriteLine("=== pipeline stages (median ms) ===");
        Console.WriteLine($"{"case",-12} {"read L",8} {"read R",8} {"src L",8} {"src R",8} {"build",9} {"markup",9} {"revrender",10} {"sum",9}");
        foreach (var r in rows)
        {
            var s = r.Stages!;
            var sum = s.ReadLeft.Median + s.ReadRight.Median + s.ReadSrcLeft.Median + s.ReadSrcRight.Median + s.Build.Median + s.Render.Median;
            Console.WriteLine($"{r.Name,-12} {s.ReadLeft.Median,8:F1} {s.ReadRight.Median,8:F1} {s.ReadSrcLeft.Median,8:F1} {s.ReadSrcRight.Median,8:F1} {s.Build.Median,9:F1} {s.Render.Median,9:F1} {s.RevRender.Median,10:F1} {sum,9:F1}");
        }
        Console.WriteLine();
        Console.WriteLine("note: 'markup' performs its own two RetainSources reads internally, so 'sum' double-counts");
        Console.WriteLine("      'src L'/'src R' and is an upper bound rather than the end-to-end figure.");
    }
}

static string Sha(byte[] b) => Convert.ToHexString(SHA256.HashData(b));

static string RevisionsFingerprint(IReadOnlyList<DocxDiffRevision> revs)
{
    var sb = new StringBuilder();
    foreach (var r in revs)
    {
        sb.Append(r.Type).Append('|').Append(r.Author).Append('|')
          .Append(r.Text).Append('|').Append(r.LeftAnchor).Append('|')
          .Append(r.RightAnchor).Append('|').Append(r.MoveGroupId).Append('|')
          .Append(r.IsMoveSource).Append('|');
        if (r.FormatChange is { } fc)
        {
            sb.Append(fc.Scope).Append('~').Append(string.Join(',', fc.ChangedPropertyNames)).Append('~')
              .Append(string.Join(',', fc.OldProperties.OrderBy(kv => kv.Key, StringComparer.Ordinal).Select(kv => $"{kv.Key}={kv.Value}"))).Append('~')
              .Append(string.Join(',', fc.NewProperties.OrderBy(kv => kv.Key, StringComparer.Ordinal).Select(kv => $"{kv.Key}={kv.Value}")));
        }

        sb.Append('\n');
    }

    return sb.ToString();
}

int IntArg(string name, int fallback)
{
    var i = Array.IndexOf(args, name);
    return i >= 0 && i + 1 < args.Length && int.TryParse(args[i + 1], out var v) ? v : fallback;
}

string? StrArg(string name)
{
    var i = Array.IndexOf(args, name);
    return i >= 0 && i + 1 < args.Length ? args[i + 1] : null;
}

sealed record Stat(double Min, double Median, double Max, double Mean, double AllocBytes);

sealed record StageTimes(Stat ReadLeft, Stat ReadRight, Stat ReadSrcLeft, Stat ReadSrcRight, Stat Build, Stat Render, Stat RevRender);

sealed class Row(string name, Stat compare)
{
    public string Name { get; } = name;
    public Stat Compare { get; } = compare;
    public Stat? Revisions { get; set; }
    public Stat? EditScript { get; set; }
    public Stat? AllProducts { get; set; }
    public StageTimes? Stages { get; set; }
}
