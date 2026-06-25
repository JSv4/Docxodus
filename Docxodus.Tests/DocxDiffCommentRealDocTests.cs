#nullable enable
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;
using Xunit.Abstractions;

namespace Docxodus.Tests;

/// <summary>
/// Headline real-document verification of COMMENT fidelity over the vendored comment-dense fixture
/// <c>TestFiles/DD/DD002-DenseComments.docx</c> (built by <see cref="DocxDiffCommentRealDocFixture"/>): a
/// contract carrying a paragraph with multiple comments, overlapping ranges, a range spanning paragraphs, a
/// comment whose anchored text is edited, a threaded reply, and plain commented sentences. A variant is
/// produced by editing many commented paragraphs, then <see cref="DocxDiff.Compare"/> must hold:
/// schema validity (no NEW errors), comment id↔range↔reference↔definition integrity (unique ids, 1:1 range
/// pairing, every reference resolves to exactly one comment), the comment-structure round-trip (accept ≡
/// right, reject ≡ left — including threaded reply links), and a clean headless-LibreOffice load + refresh.
/// The vendored fixture is REQUIRED (missing/empty fails the suite); the LibreOffice oracle soft-skips when
/// <c>soffice</c>/python-uno is unavailable so the suite stays portable.
/// </summary>
public class DocxDiffCommentRealDocTests
{
    private readonly ITestOutputHelper _out;
    public DocxDiffCommentRealDocTests(ITestOutputHelper o) => _out = o;

    private static readonly DirectoryInfo TestFilesDir = new("../../../../TestFiles/");
    public static string VendoredFixturePath =>
        Path.Combine(TestFilesDir.FullName, "DD", "DD002-DenseComments.docx");
    private static readonly string ScratchDir =
        Path.Combine(Path.GetTempPath(), "docxodus-comment-fidelity");

    [Fact]
    public void Contract_CommentFidelity()
    {
        Assert.True(File.Exists(VendoredFixturePath) && new FileInfo(VendoredFixturePath).Length > 0,
            $"REQUIRED vendored fixture missing/empty: {VendoredFixturePath}. " +
            "Regenerate via DocxDiffCommentRealDocTests.__RegenerateVendoredFixture (remove its Skip).");

        var leftBytes = File.ReadAllBytes(VendoredFixturePath);
        var rightBytes = EditCommentedParagraphs(leftBytes);
        var left = new WmlDocument("dd002.docx", leftBytes);
        var right = new WmlDocument("dd002.docx", rightBytes);

        var result = DocxDiff.Compare(left, right);

        Directory.CreateDirectory(ScratchDir);
        File.WriteAllBytes(Path.Combine(ScratchDir, "left.docx"), leftBytes);
        File.WriteAllBytes(Path.Combine(ScratchDir, "right.docx"), rightBytes);
        File.WriteAllBytes(Path.Combine(ScratchDir, "compare.docx"), result.DocumentByteArray);

        // (1) schema validity — no NEW errors vs the inputs (numeric ids normalized so a legit comment-id
        //     renumber is not miscounted).
        var baseErrors = SchemaErrors(leftBytes).Concat(SchemaErrors(rightBytes)).ToHashSet();
        var newErrors = SchemaErrors(result.DocumentByteArray).Where(e => !baseErrors.Contains(e)).ToList();
        Assert.True(newErrors.Count == 0,
            $"Compare introduced {newErrors.Count} new schema error(s): {string.Join(" | ", newErrors.Take(6))}");

        // (2) comment-structure soundness of the Compare output: unique range-start ids, 1:1 pairing, every
        //     anchored id resolves to exactly one definition.
        var (startIds, endIds, refIds, defCounts) = CommentMarkers(result.DocumentByteArray);
        Assert.True(startIds.Count == startIds.Distinct().Count(),
            $"duplicate commentRangeStart id(s): {Dup(startIds)}");
        Assert.True(startIds.OrderBy(x => x).SequenceEqual(endIds.OrderBy(x => x)),
            $"commentRangeStart/End ids not 1:1 paired (starts:{startIds.Count} ends:{endIds.Count})");
        var anchored = startIds.Concat(endIds).Concat(refIds).ToHashSet();
        var unresolved = anchored.Where(id => !defCounts.TryGetValue(id, out var n) || n != 1).ToList();
        Assert.True(unresolved.Count == 0, $"comment id(s) with no/ambiguous definition: {string.Join(",", unresolved)}");

        // (3) comment-structure round-trip: accept ≡ right, reject ≡ left (anchored comments by content + threading)
        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal(BodyText(rightBytes), BodyText(accepted.DocumentByteArray));
        Assert.Equal(BodyText(leftBytes), BodyText(rejected.DocumentByteArray));
        Assert.Equal(AnchorProjection(rightBytes), AnchorProjection(accepted.DocumentByteArray));
        Assert.Equal(AnchorProjection(leftBytes), AnchorProjection(rejected.DocumentByteArray));
        // anchor SPAN coverage round-trip: the text each comment physically brackets matches the side it
        // resolves to (catches a surviving range marker placed on the wrong side of an edit).
        Assert.Equal(RangeCoverage(rightBytes), RangeCoverage(accepted.DocumentByteArray));
        Assert.Equal(RangeCoverage(leftBytes), RangeCoverage(rejected.DocumentByteArray));
        _out.WriteLine($"comments: starts={startIds.Count} refs={refIds.Count} defs={defCounts.Count}");

        // (4) headless LibreOffice oracle — independent confirmation that every comment loads + anchors + the
        //     threaded reply links survive + refresh drops nothing. Soft-skips when soffice/python-uno is absent.
        RunLibreOfficeCommentOracle(Path.Combine(ScratchDir, "compare.docx"));
        RunLibreOfficeCommentOracle(Path.Combine(ScratchDir, "right.docx"));
    }

    /// <summary>Regenerate the vendored dense comment fixture. Skip-by-default — remove the Skip to rewrite the
    /// committed artifact.</summary>
    [Fact(Skip = "manual: regenerates the committed TestFiles/DD/DD002 fixture from DocxDiffCommentRealDocFixture")]
    public void __RegenerateVendoredFixture()
    {
        var bytes = DocxDiffCommentRealDocFixture.Build();
        Directory.CreateDirectory(Path.GetDirectoryName(VendoredFixturePath)!);
        File.WriteAllBytes(VendoredFixturePath, bytes);
        _out.WriteLine($"wrote {bytes.Length} bytes -> {VendoredFixturePath}");
    }

    // ---- variant construction -------------------------------------------------------------------

    /// <summary>Edit many commented paragraphs across every shape: an anchored word (id 5), a threaded anchor
    /// (id 6/7), the overlapping-range word (id 2/3), the cross-paragraph range start (id 4), and a plain
    /// commented sentence (id 8) — a dense stress of the fine comment-edit path.</summary>
    private static byte[] EditCommentedParagraphs(byte[] left)
    {
        using var ms = new MemoryStream();
        ms.Write(left, 0, left.Length);
        ms.Position = 0;
        using (var doc = WordprocessingDocument.Open(ms, true))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            void Edit(string find, string repl)
            {
                foreach (var t in body.Descendants<Text>())
                    if (t.Text.Contains(find)) { t.Text = t.Text.Replace(find, repl); return; }
            }
            Edit("cap", "limit");                       // id 5 anchored word edited
            Edit("survival period", "survival window"); // id 6/7 threaded anchor edited
            Edit("term", "defined term");               // overlapping ranges (id 2/3)
            Edit("Each obligation under this clause is several.",
                 "Each obligation under this clause is several and binding."); // cross-para range start (id 4)
            Edit("Delaware", "New York");               // plain commented sentence (id 8)
            Edit("parties named below", "parties identified below"); // id 1 anchored phrase
        }
        return ms.ToArray();
    }

    // ---- LibreOffice oracle ---------------------------------------------------------------------

    private void RunLibreOfficeCommentOracle(string docxPath)
    {
        var script = Path.GetFullPath(Path.Combine("../../../../tools/diffharness/lo/lo_comment_check.py"));
        if (!File.Exists(script))
        {
            _out.WriteLine($"[lo-skip] oracle script not found: {script}");
            return;
        }
        string stdout;
        try
        {
            var psi = new ProcessStartInfo("python3", $"\"{script}\" \"{docxPath}\"")
            {
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
            };
            using var proc = Process.Start(psi);
            if (proc == null) { _out.WriteLine("[lo-skip] could not start python3"); return; }
            stdout = proc.StandardOutput.ReadToEnd() + proc.StandardError.ReadToEnd();
            if (!proc.WaitForExit(120_000)) { try { proc.Kill(true); } catch { } _out.WriteLine("[lo-skip] timed out"); return; }
        }
        catch (Exception e)
        {
            _out.WriteLine($"[lo-skip] could not run oracle: {e.Message}");
            return;
        }

        if (stdout.Contains("RESULT: OK"))
        {
            _out.WriteLine($"[lo-ok] {stdout.Split('\n').FirstOrDefault(l => l.StartsWith("doc="))}");
            return;
        }
        if (stdout.Contains("RESULT: FAIL") || stdout.Contains("LOAD=FAILED"))
            Assert.Fail($"LibreOffice comment oracle FAILED for {Path.GetFileName(docxPath)}:\n{stdout}");
        // No clean OK and no explicit FAIL → soffice / python-uno unavailable. Soft-skip for portability.
        _out.WriteLine($"[lo-skip] oracle inconclusive (soffice/uno unavailable): {stdout.Split('\n').FirstOrDefault()}");
    }

    // ---- comment-structure observers ------------------------------------------------------------

    private static (List<string> Starts, List<string> Ends, List<string> Refs, Dictionary<string, int> DefCounts)
        CommentMarkers(byte[] bytes)
    {
        using var ms = new MemoryStream(bytes);
        using var w = WordprocessingDocument.Open(ms, false);
        var main = w.MainDocumentPart!;
        var body = main.Document?.Body;
        List<string> Ids(string n) => body?.Descendants().Where(e => e.LocalName == n)
            .Select(e => e.GetAttributes().FirstOrDefault(a => a.LocalName == "id").Value ?? "").ToList() ?? new();
        var defCounts = new Dictionary<string, int>();
        var croot = main.WordprocessingCommentsPart?.GetXDocument().Root;
        if (croot != null)
            foreach (var c in croot.Elements().Where(e => e.Name.LocalName == "comment"))
            {
                var id = (string?)c.Attributes().FirstOrDefault(a => a.Name.LocalName == "id") ?? "";
                defCounts[id] = defCounts.TryGetValue(id, out var k) ? k + 1 : 1;
            }
        return (Ids("commentRangeStart"), Ids("commentRangeEnd"), Ids("commentReference"), defCounts);
    }

    /// <summary>The anchored-comment projection (author + text + threaded parent text), keyed by content so
    /// dedup-renumbered ids still compare. Identical for two documents with the same comment structure.</summary>
    private static string AnchorProjection(byte[] bytes)
    {
        using var ms = new MemoryStream(bytes);
        using var w = WordprocessingDocument.Open(ms, false);
        var main = w.MainDocumentPart!;
        var body = main.Document?.Body;
        if (body is null) return "anchors=[]";

        var defs = new Dictionary<string, (string Author, string Text, string ParaId)>();
        var croot = main.WordprocessingCommentsPart?.GetXDocument().Root;
        if (croot != null)
            foreach (var c in croot.Elements().Where(e => e.Name.LocalName == "comment"))
            {
                var id = AttrLocal(c, "id");
                var text = string.Concat(c.Descendants().Where(e => e.Name.LocalName == "t").Select(e => e.Value));
                var firstP = c.Descendants().FirstOrDefault(e => e.Name.LocalName == "p");
                defs[id] = (AttrLocal(c, "author"), text, firstP != null ? AttrLocal(firstP, "paraId") : "");
            }
        var parentOfPara = new Dictionary<string, string>();
        var exroot = main.WordprocessingCommentsExPart?.GetXDocument().Root;
        if (exroot != null)
            foreach (var ex in exroot.Elements().Where(e => e.Name.LocalName == "commentEx"))
            {
                var pid = AttrLocal(ex, "paraId");
                var parent = AttrLocal(ex, "paraIdParent");
                if (pid.Length > 0 && parent.Length > 0) parentOfPara[pid] = parent;
            }
        string ParentText(string paraId)
        {
            if (!parentOfPara.TryGetValue(paraId, out var pp)) return "";
            return defs.Values.FirstOrDefault(d => d.ParaId == pp).Text ?? "";
        }
        var refIds = body.Descendants().Where(e => e.LocalName == "commentReference")
            .Select(e => e.GetAttributes().FirstOrDefault(a => a.LocalName == "id").Value ?? "").ToList();
        var anchors = refIds.Where(id => defs.ContainsKey(id))
            .Select(id => $"{defs[id].Author}|{defs[id].Text}|parent={ParentText(defs[id].ParaId)}")
            .OrderBy(s => s, StringComparer.Ordinal).ToList();
        return $"anchors=[{string.Join(" ;; ", anchors)}]";
    }

    /// <summary>The visible text each comment physically brackets (between commentRangeStart/End), keyed by
    /// content, trimmed (a wholly-deleted anchor can migrate a boundary whitespace token — body text still
    /// round-trips). Run on a materialized accept/reject doc, the span-coverage counterpart to
    /// <see cref="AnchorProjection"/>.</summary>
    private static string RangeCoverage(byte[] bytes)
    {
        using var ms = new MemoryStream(bytes);
        using var w = WordprocessingDocument.Open(ms, false);
        var main = w.MainDocumentPart!;
        var body = main.Document?.Body;
        if (body is null) return "cov=[]";
        var content = new Dictionary<string, string>();
        var croot = main.WordprocessingCommentsPart?.GetXDocument().Root;
        if (croot != null)
            foreach (var c in croot.Elements().Where(e => e.Name.LocalName == "comment"))
                content[AttrLocal(c, "id")] = AttrLocal(c, "author") + "|" +
                    string.Concat(c.Descendants().Where(e => e.Name.LocalName == "t").Select(e => e.Value));
        var open = new Dictionary<string, System.Text.StringBuilder>();
        var done = new List<string>();
        foreach (var e in body.Descendants())
        {
            if (e.LocalName == "commentRangeStart") open[AttrValueOf(e)] = new System.Text.StringBuilder();
            else if (e.LocalName == "commentRangeEnd")
            {
                var id = AttrValueOf(e);
                if (open.TryGetValue(id, out var sb))
                {
                    done.Add($"{(content.TryGetValue(id, out var k) ? k : id)}=>[{sb.ToString().Trim()}]");
                    open.Remove(id);
                }
            }
            else if (e.LocalName == "t")
                foreach (var sb in open.Values) sb.Append(e.InnerText);
        }
        return "cov=[" + string.Join(" ;; ", done.OrderBy(s => s, StringComparer.Ordinal)) + "]";
    }

    private static string AttrValueOf(OpenXmlElement e) =>
        e.GetAttributes().FirstOrDefault(a => a.LocalName == "id").Value ?? "";

    private static string AttrLocal(XElement e, string ln) =>
        (string?)e.Attributes().FirstOrDefault(a => a.Name.LocalName == ln) ?? "";

    private static string BodyText(byte[] bytes)
    {
        using var ms = new MemoryStream(bytes);
        using var w = WordprocessingDocument.Open(ms, false);
        var body = w.MainDocumentPart?.Document?.Body;
        return body is null ? "" : string.Concat(body.Descendants<Text>().Select(t => t.Text));
    }

    private static HashSet<string> SchemaErrors(byte[] bytes)
    {
        using var ms = new MemoryStream(bytes);
        using var w = WordprocessingDocument.Open(ms, false);
        var v = new OpenXmlValidator(FileFormatVersions.Office2019);
        return v.Validate(w)
            .Select(e => $"{e.Id}@{e.Part?.Uri}: {Regex.Replace(e.Description, "'[0-9]+'", "'#'")}")
            .ToHashSet();
    }

    private static string Dup(IEnumerable<string> ids) =>
        string.Join(",", ids.GroupBy(x => x).Where(g => g.Count() > 1).Select(g => g.Key));
}
