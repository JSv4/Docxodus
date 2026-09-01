// Complex-form-document benchmark harness.
//
// Exercises the surfaces an agentic caller leans on hardest — markdown/HTML projection,
// tracked-change session editing, DocxDiff redlining — against a single
// heavyweight .docx and verifies the invariants that matter for legal work:
//
//   * no-edit open→save round-trip is text-exact and part-preserving
//   * accept-all on a redline reproduces the revised document's text exactly
//   * reject-all on a redline reproduces the baseline document's text exactly
//   * outputs add zero OOXML schema findings over the source document's own baseline
//
// The reference document class is an NVCA-style model legal form (dense footnotes,
// hundreds of bookmarks and cross-reference fields, multilevel numbering, multiple
// sections and header/footer variants). The document itself is not committed — pass any
// comparable .docx on the command line. See README.md and FINDINGS.md.

using System.Diagnostics;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus;

if (args.Length < 1)
{
    Console.Error.WriteLine("usage: ComplexFormBenchmark <document.docx> [edits.json] [--out <dir>]");
    Console.Error.WriteLine("       edits.json defaults to edits/nvca-coi.json next to this program");
    return 1;
}

var docPath = args[0];
var editsPath = args.Length > 1 && !args[1].StartsWith("--")
    ? args[1]
    : Path.Combine(AppContext.BaseDirectory, "edits", "nvca-coi.json");
var outIx = Array.IndexOf(args, "--out");
var outDir = outIx >= 0 && outIx + 1 < args.Length ? args[outIx + 1] : Path.Combine(Path.GetTempPath(), "complex-form-benchmark");
Directory.CreateDirectory(outDir);

var bytes = File.ReadAllBytes(docPath);
var edits = JsonSerializer.Deserialize<EditScript>(File.ReadAllText(editsPath), new JsonSerializerOptions { PropertyNameCaseInsensitive = true })!;
Console.WriteLine($"document: {docPath} ({bytes.Length:N0} bytes)");
Console.WriteLine($"edit script: {editsPath} ({edits.Replacements.Count} replacements + delete/insert/format/comment)");
Console.WriteLine($"outputs: {outDir}");
Console.WriteLine();

var baselineErrors = ValidateCount(bytes);
Console.WriteLine($"source schema findings baseline: {baselineErrors}");

// ---------- 1. Projections ----------
var doc = new WmlDocument(Path.GetFileName(docPath), bytes);
Bench("html (footnotes+headers rendered)", () =>
{
    var html = WmlToHtmlConverter.ConvertToHtml(doc, new WmlToHtmlConverterSettings
    {
        FabricateCssClasses = true,
        RenderFootnotesAndEndnotes = true,
        RenderHeadersAndFooters = true,
    });
    File.WriteAllText(Path.Combine(outDir, "document.html"), WmlToHtmlConverter.ToHtmlString(html, indent: false));
});
Bench("markdown projection", () =>
{
    var projection = WmlToMarkdownConverter.Convert(doc, new WmlToMarkdownConverterSettings());
    File.WriteAllText(Path.Combine(outDir, "document.md"), projection.Markdown);
    Console.WriteLine($"    {projection.Markdown.Length:N0} chars, {projection.AnchorIndex.Count} anchors");
});
Bench("DocxDiff compatibility probe", () =>
{
    var report = DocxDiff.InspectCompatibility(doc);
    Console.WriteLine($"    warnings: {report.Warnings.Count}");
});

// ---------- 2. No-edit round-trip ----------
Bench("no-edit session round-trip", () =>
{
    byte[] rt;
    using (var s = new DocxSession(bytes)) rt = s.Save();
    Check("round-trip text exact", ExtractText(rt) == ExtractText(bytes));
    Check("round-trip parts identical", SamePartNames(bytes, rt));
});

// ---------- 3. Tracked-change session (the agentic redline path) ----------
byte[] trackedDocx = [];
Bench("tracked session: full edit script", () =>
{
    using var s = new DocxSession(bytes, new DocxSessionSettings { RevisionAuthor = edits.Author });
    s.SetTrackedChanges(TrackedChangeMode.RenderInline);
    var applied = ApplyEdits(s, edits, out var failures);
    Console.WriteLine($"    edits applied: {applied}, failed: {failures}");
    var revs = s.ListRevisions();
    Console.WriteLine($"    native revisions recorded: {revs.Count} (author \"{edits.Author}\")");
    if (revs.Count >= 2)
    {
        Check("accept single revision", s.AcceptRevision(revs[0].Id).Success);
        Check("reject single revision", s.RejectRevision(revs[^1].Id).Success);
    }
    trackedDocx = s.Save();
    File.WriteAllBytes(Path.Combine(outDir, "tracked-session.docx"), trackedDocx);
    Check("tracked save adds no schema findings", ValidateCount(trackedDocx) <= baselineErrors);
});

// ---------- 4. Clean edits + DocxDiff redline ----------
byte[] modified = [];
Bench("clean session: same edit script untracked", () =>
{
    using var s = new DocxSession(bytes);
    ApplyEdits(s, edits, out _);
    modified = s.Save();
    File.WriteAllBytes(Path.Combine(outDir, "modified.docx"), modified);
});

var left = new WmlDocument("baseline.docx", bytes);
var right = new WmlDocument("modified.docx", modified);
WmlDocument? redline = null;
Bench("DocxDiff.Compare", () =>
{
    redline = DocxDiff.Compare(left, right);
    File.WriteAllBytes(Path.Combine(outDir, "redline-docxdiff.docx"), redline.DocumentByteArray);
});
Bench("DocxDiff.GetRevisions", () =>
{
    Console.WriteLine($"    revisions: {DocxDiff.GetRevisions(left, right).Count}");
});
Bench("DocxDiff edit script + semantic changes", () =>
{
    var editScript = DocxDiff.GetEditScriptJson(left, right);
    var semantic = DocxDiff.GetSemanticChangesJson(left, right);
    File.WriteAllText(Path.Combine(outDir, "editscript.json"), editScript);
    File.WriteAllText(Path.Combine(outDir, "semantic-changes.json"), semantic);
    Console.WriteLine($"    edit script {editScript.Length:N0} chars, semantic changes {semantic.Length:N0} chars");
});
Bench("DocxDiff round-trip invariants", () =>
{
    var acceptAll = RevisionProcessor.AcceptRevisions(redline!);
    var rejectAll = RevisionProcessor.RejectRevisions(redline!);
    Check("accept-all == modified text", ExtractText(acceptAll.DocumentByteArray) == ExtractText(modified));
    Check("reject-all == baseline text", ExtractText(rejectAll.DocumentByteArray) == ExtractText(bytes));
    Check("redline adds no schema findings", ValidateCount(redline!.DocumentByteArray) <= baselineErrors);
});
Bench("redline -> HTML with tracked-change markup", () =>
{
    var html = WmlToHtmlConverter.ConvertToHtml(redline!, new WmlToHtmlConverterSettings
    {
        FabricateCssClasses = true,
        RenderTrackedChanges = true,
        RenderFootnotesAndEndnotes = true,
    });
    File.WriteAllText(Path.Combine(outDir, "redline.html"), WmlToHtmlConverter.ToHtmlString(html, indent: false));
});

// ---------- helpers ----------

static void Bench(string label, Action act)
{
    var sw = Stopwatch.StartNew();
    try
    {
        act();
        sw.Stop();
        Console.WriteLine($"[bench] {label}: {sw.ElapsedMilliseconds} ms");
    }
    catch (Exception ex)
    {
        sw.Stop();
        Console.WriteLine($"[bench] {label}: FAILED after {sw.ElapsedMilliseconds} ms :: {ex.GetType().Name}: {ex.Message}");
        FailedChecks++;
    }
}

static int ApplyEdits(DocxSession s, EditScript edits, out int failures)
{
    int ok = 0, bad = 0;
    void Run(string what, Func<EditResult?> op)
    {
        var r = op();
        if (r is { Success: true }) ok++;
        else { bad++; Console.WriteLine($"    [edit failed] {what}: {r?.Error?.Code.ToString() ?? "no match"}"); }
    }

    foreach (var rep in edits.Replacements)
        Run($"replace \"{rep.Find}\"", () =>
        {
            var m = s.Grep(System.Text.RegularExpressions.Regex.Escape(rep.Find)).FirstOrDefault();
            return m is null ? null : s.ReplaceMatch(m, rep.Replace);
        });

    if (edits.DeleteBlockContaining is { } del)
        Run("delete block", () =>
        {
            var m = s.Grep(System.Text.RegularExpressions.Regex.Escape(del)).FirstOrDefault();
            return m is null ? null : s.DeleteBlock(m.EnclosingAnchor.Anchor.Id);
        });

    if (edits.InsertAfterBlockContaining is { } ins)
        Run("insert paragraph", () =>
        {
            var m = s.Grep(System.Text.RegularExpressions.Regex.Escape(ins.Needle)).FirstOrDefault();
            return m is null ? null : s.InsertParagraph(m.EnclosingAnchor.Anchor.Id, Position.After, ins.Markdown);
        });

    if (edits.Italicize is { } ital)
        Run("italicize span", () =>
        {
            var m = s.Grep(System.Text.RegularExpressions.Regex.Escape(ital)).FirstOrDefault();
            return m is null ? null : s.ApplyFormat(m, new FormatOp { Italic = true });
        });

    if (edits.Comment is { } com)
        Run("add comment", () =>
        {
            var m = s.Grep(System.Text.RegularExpressions.Regex.Escape(com.Needle)).FirstOrDefault();
            return m is null ? null : s.AddComment(m.EnclosingAnchor.Anchor.Id, m.Span, edits.Author, com.Text);
        });

    failures = bad;
    return ok;
}

static string ExtractText(byte[] docx)
{
    XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    var sb = new StringBuilder();
    using var ms = new MemoryStream(docx);
    using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
    foreach (var partName in new[] { "word/document.xml", "word/footnotes.xml", "word/endnotes.xml" })
    {
        var entry = zip.GetEntry(partName);
        if (entry is null) continue;
        var xdoc = XDocument.Load(entry.Open());
        foreach (var p in xdoc.Descendants(w + "p"))
        {
            foreach (var t in p.Descendants())
            {
                if (t.Name == w + "t") sb.Append(t.Value);
                else if (t.Name == w + "tab") sb.Append('\t');
            }
            sb.Append('\n');
        }
    }
    return sb.ToString();
}

static bool SamePartNames(byte[] a, byte[] b)
{
    static HashSet<string> Names(byte[] docx)
    {
        using var ms = new MemoryStream(docx);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        return zip.Entries.Select(e => e.FullName).ToHashSet();
    }
    return Names(a).SetEquals(Names(b));
}

static int ValidateCount(byte[] docx)
{
    using var ms = new MemoryStream(docx);
    using var word = WordprocessingDocument.Open(ms, false);
    return new OpenXmlValidator().Validate(word).Count();
}

static void Check(string label, bool pass)
{
    Console.WriteLine($"    [check] {label}: {(pass ? "PASS" : "FAIL")}");
    if (!pass) FailedChecks++;
}

internal sealed record EditScript
{
    public string Author { get; init; } = "Benchmark Reviewer";
    public List<Replacement> Replacements { get; init; } = [];
    public string? DeleteBlockContaining { get; init; }
    public InsertSpec? InsertAfterBlockContaining { get; init; }
    public string? Italicize { get; init; }
    public CommentSpec? Comment { get; init; }
}

internal sealed record Replacement(string Find, string Replace);
internal sealed record InsertSpec(string Needle, string Markdown);
internal sealed record CommentSpec(string Needle, string Text);

internal static partial class Program
{
    internal static int FailedChecks;
}
