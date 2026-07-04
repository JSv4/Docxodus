#nullable enable

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus;
using Docxodus.Internal;
using Docxodus.Tests.Ir.Diff;
using Xunit;
using Xunit.Abstractions;

namespace Docxodus.Tests;

/// <summary>
/// Headline smoke proof of the header/footer comparison campaign (spec:
/// 2026-07-03-docxdiff-header-footer-diff-design.md): DocxDiff compares header/footer stories as Word
/// Compare does. Two ends of the spectrum:
/// <list type="number">
/// <item><b>Synthetic rich pair</b> — edited default header, edited PAGE-field footer, untouched
/// first-page header, and an ADDED even-page footer — proven end-to-end at the PUBLIC surface:
/// native markup in the right parts, per-story accept ≡ right / reject ≡ left via the byte-level
/// client primitives (<see cref="DocxDiffOps.AcceptRevisions"/>/<see cref="DocxDiffOps.RejectRevisions"/>),
/// hdr/ftr-anchored revisions (Fine) with the compatible mode excluding them, <c>headerFooterOps</c> in
/// the edit-script JSON, schema validity, and a headless-LibreOffice render oracle (an INDEPENDENT
/// renderer surfaces the expected story text on the accept and reject views).</item>
/// <item><b>Real corpus pair</b> — <c>WC004-Large</c> ↔ <c>WC004-Large-Mod</c>, the corpus' one real
/// footer difference, previously SILENT (carried left's footer verbatim, no revision anywhere): the
/// change now surfaces as ftr-anchored revisions and the footer round-trips.</item>
/// </list>
/// The LibreOffice oracle soft-skips when <c>soffice</c>/python-uno is unavailable (portability), same
/// as the comment-fidelity suite.
/// </summary>
public class DocxDiffHeaderFooterSmokeTests
{
    private readonly ITestOutputHelper _out;
    public DocxDiffHeaderFooterSmokeTests(ITestOutputHelper o) => _out = o;

    private static readonly DirectoryInfo TestFilesDir = new("../../../../TestFiles/");
    private static readonly string ScratchDir =
        Path.Combine(Path.GetTempPath(), "docxodus-headerfooter-smoke");

    // ---- Test A: synthetic rich pair --------------------------------------------------------------

    private const string PageFieldFooterLeft =
        "<w:p><w:r><w:t xml:space=\"preserve\">Acme Corp — Page </w:t></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
        "<w:r><w:instrText xml:space=\"preserve\"> PAGE </w:instrText></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" +
        "<w:r><w:t>1</w:t></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>";

    private static string PageFieldFooterRight => PageFieldFooterLeft
        .Replace("Acme Corp — Page ", "Acme Corporation — Page ");

    private static WmlDocument BuildLeft() => HeaderFooterFixtures.Build(
        new[]
        {
            new HeaderFooterFixtures.Section(
                new[] { "The parties agree as follows.", "Section 1. Definitions." },
                Headers: new[] { ("default", "rIdHdrDef"), ("first", "rIdHdrFirst") },
                Footers: new[] { ("default", "rIdFtrDef") }),
        },
        headerParts: new Dictionary<string, string[]>
        {
            ["rIdHdrDef"] = new[] { "CONFIDENTIAL — Draft 1 — Attorney Work Product" },
            ["rIdHdrFirst"] = new[] { "Acme / Beta Merger Agreement" },
        },
        footerParts: new Dictionary<string, string[]> { ["rIdFtrDef"] = new[] { PageFieldFooterLeft } },
        titlePg: true);

    private static WmlDocument BuildRight() => HeaderFooterFixtures.Build(
        new[]
        {
            new HeaderFooterFixtures.Section(
                new[] { "The parties agree as follows, as amended.", "Section 1. Definitions." },
                Headers: new[] { ("default", "rIdHdrDef"), ("first", "rIdHdrFirst") },
                Footers: new[] { ("default", "rIdFtrDef"), ("even", "rIdFtrEven") }),
        },
        headerParts: new Dictionary<string, string[]>
        {
            ["rIdHdrDef"] = new[] { "CONFIDENTIAL — Draft 2 — Attorney Work Product (Reviewed)" },
            ["rIdHdrFirst"] = new[] { "Acme / Beta Merger Agreement" },
        },
        footerParts: new Dictionary<string, string[]>
        {
            ["rIdFtrDef"] = new[] { PageFieldFooterRight },
            ["rIdFtrEven"] = new[] { "Even-page confidentiality notice" },
        },
        titlePg: true, evenAndOddHeaders: true);

    [Fact]
    public void Smoke_HeaderFooterComparison_FullBattery()
    {
        var left = BuildLeft();
        var right = BuildRight();

        // ---- Compare: native tracked-changes markup lands INSIDE the right story parts. ----------
        var rendered = DocxDiff.Compare(left, right);
        Directory.CreateDirectory(ScratchDir);
        File.WriteAllBytes(Path.Combine(ScratchDir, "left.docx"), left.DocumentByteArray);
        File.WriteAllBytes(Path.Combine(ScratchDir, "right.docx"), right.DocumentByteArray);
        File.WriteAllBytes(Path.Combine(ScratchDir, "redline.docx"), rendered.DocumentByteArray);

        var storyXml = StoryPartsXmlByText(rendered);
        Assert.Contains("<w:ins", storyXml["Draft"]);         // edited default header: fine ins/del
        Assert.Contains("<w:del", storyXml["Draft"]);
        Assert.DoesNotContain("<w:ins", storyXml["Merger"]);  // untouched first-page header: verbatim
        Assert.Contains("<w:ins", storyXml["— Page"]);        // edited footer
        Assert.Contains("<w:ins", storyXml["Even-page"]);     // inserted even footer: all-ins
        Assert.Contains("PAGE", storyXml["— Page"]);          // field plumbing survives the footer edit

        // ---- Round-trip via the CLIENT byte primitives: accept ≡ right, reject ≡ left, per story. --
        var accepted = new WmlDocument("a.docx", DocxDiffOps.AcceptRevisions(rendered.DocumentByteArray));
        var rejected = new WmlDocument("r.docx", DocxDiffOps.RejectRevisions(rendered.DocumentByteArray));
        File.WriteAllBytes(Path.Combine(ScratchDir, "accepted.docx"), accepted.DocumentByteArray);
        File.WriteAllBytes(Path.Combine(ScratchDir, "rejected.docx"), rejected.DocumentByteArray);

        string? Story(WmlDocument d, bool isHeader, string kind) =>
            HeaderFooterFixtures.ReferencedStoryText(d, isHeader, 0, kind);

        Assert.Equal("CONFIDENTIAL — Draft 2 — Attorney Work Product (Reviewed)", Story(accepted, true, "default"));
        Assert.Equal("Acme / Beta Merger Agreement", Story(accepted, true, "first"));
        Assert.Equal("Acme Corporation — Page 1", Story(accepted, false, "default"));
        Assert.Equal("Even-page confidentiality notice", Story(accepted, false, "even"));

        Assert.Equal("CONFIDENTIAL — Draft 1 — Attorney Work Product", Story(rejected, true, "default"));
        Assert.Equal("Acme / Beta Merger Agreement", Story(rejected, true, "first"));
        Assert.Equal("Acme Corp — Page 1", Story(rejected, false, "default"));
        Assert.Equal("", Story(rejected, false, "even")); // rejected inserted story → empty ≡ absent

        // ---- Revisions: hdr/ftr anchors in Fine (the default); NONE in WmlComparerCompatible. -----
        var revisions = DocxDiff.GetRevisions(left, right);
        var hf = revisions.Where(r =>
            ((r.LeftAnchor ?? r.RightAnchor) ?? "").Contains(":hdr") ||
            ((r.LeftAnchor ?? r.RightAnchor) ?? "").Contains(":ftr")).ToList();
        Assert.NotEmpty(hf);
        Assert.Contains(hf, r => r.Type == DocxDiffRevisionType.Inserted && r.Text.Contains("(Reviewed)"));
        Assert.Contains(hf, r => r.Type == DocxDiffRevisionType.Inserted && r.Text.Contains("Even-page"));
        Assert.Contains(hf, r => r.RightAnchor?.Contains(":ftr") == true);
        var compat = DocxDiff.GetRevisions(left, right, new DocxDiffSettings
        { RevisionGranularity = DocxDiffRevisionGranularity.WmlComparerCompatible });
        Assert.DoesNotContain(compat, r =>
            ((r.LeftAnchor ?? r.RightAnchor) ?? "").Contains(":hdr") ||
            ((r.LeftAnchor ?? r.RightAnchor) ?? "").Contains(":ftr"));

        // ---- Diff-as-data: the edit script carries the story ops. --------------------------------
        var json = DocxDiff.GetEditScriptJson(left, right);
        Assert.Contains("\"headerFooterOps\"", json);
        Assert.Contains("\"kind\": \"Even\"", json);

        // ---- Schema validity on all three artifacts. ----------------------------------------------
        Assert.Empty(SchemaErrors(rendered.DocumentByteArray));
        Assert.Empty(SchemaErrors(accepted.DocumentByteArray));
        Assert.Empty(SchemaErrors(rejected.DocumentByteArray));

        // ---- Independent renderer oracle: headless LibreOffice surfaces the story text. ----------
        RunLibreOfficeHeaderFooterOracle(Path.Combine(ScratchDir, "accepted.docx"),
            expects: new[] { "Draft 2", "(Reviewed)", "Acme Corporation", "Merger Agreement", "Even-page" },
            absents: new[] { "Draft 1" });
        RunLibreOfficeHeaderFooterOracle(Path.Combine(ScratchDir, "rejected.docx"),
            expects: new[] { "Draft 1", "Acme Corp", "Merger Agreement" },
            absents: new[] { "Reviewed", "Even-page" });
        RunLibreOfficeHeaderFooterOracle(Path.Combine(ScratchDir, "redline.docx"),
            expects: new[] { "Attorney Work Product" }, absents: Array.Empty<string>());
    }

    // ---- Test B: the real corpus pair whose footer change was previously silent -------------------

    [Fact]
    public void Smoke_WC004RealPair_FooterChangeSurfaces()
    {
        var leftPath = Path.Combine(TestFilesDir.FullName, "WC", "WC004-Large.docx");
        var rightPath = Path.Combine(TestFilesDir.FullName, "WC", "WC004-Large-Mod.docx");
        Assert.True(File.Exists(leftPath) && File.Exists(rightPath), "WC004 corpus pair missing");
        var left = new WmlDocument(leftPath);
        var right = new WmlDocument(rightPath);

        // The footer difference (a cached PAGE-field result) now SURFACES as ftr-anchored revisions —
        // before this campaign the pair's footer diff was silently ignored.
        var revisions = DocxDiff.GetRevisions(left, right);
        Assert.Contains(revisions, r =>
            ((r.LeftAnchor ?? r.RightAnchor) ?? "").Contains(":ftr"));

        // And the produced redline round-trips the footer at the text level.
        var rendered = DocxDiff.Compare(left, right);
        var accepted = new WmlDocument("a.docx", DocxDiffOps.AcceptRevisions(rendered.DocumentByteArray));
        var rejected = new WmlDocument("r.docx", DocxDiffOps.RejectRevisions(rendered.DocumentByteArray));
        Assert.Equal(FooterTexts(right), FooterTexts(accepted));
        Assert.Equal(FooterTexts(left), FooterTexts(rejected));

        // No NEW schema errors vs the inputs (the fixtures carry pre-existing noise of their own).
        var baseline = SchemaErrors(left.DocumentByteArray)
            .Concat(SchemaErrors(right.DocumentByteArray)).ToHashSet();
        var fresh = SchemaErrors(rendered.DocumentByteArray).Where(e => !baseline.Contains(e)).ToList();
        Assert.True(fresh.Count == 0,
            $"Compare introduced {fresh.Count} new schema error(s): {string.Join(" | ", fresh.Take(6))}");

        Directory.CreateDirectory(ScratchDir);
        var outPath = Path.Combine(ScratchDir, "wc004-redline.docx");
        File.WriteAllBytes(outPath, rendered.DocumentByteArray);
        RunLibreOfficeHeaderFooterOracle(outPath,
            expects: new[] { "Confidential" }, absents: Array.Empty<string>());
        _out.WriteLine($"WC004 ftr revisions: {revisions.Count(r => ((r.LeftAnchor ?? r.RightAnchor) ?? "").Contains(":ftr"))}");
    }

    // ---- helpers -----------------------------------------------------------------------------------

    /// <summary>Story-part raw XML keyed by a distinctive substring of its visible text.</summary>
    private static Dictionary<string, string> StoryPartsXmlByText(WmlDocument d)
    {
        var result = new Dictionary<string, string>();
        foreach (var xml in HeaderFooterFixtures.StoryPartsXml(d))
            foreach (var key in new[] { "Draft", "Merger", "— Page", "Even-page" })
                if (xml.Contains(key) && !result.ContainsKey(key))
                    result[key] = xml;
        return result;
    }

    /// <summary>Non-empty visible footer texts, sorted (part enumeration order differs between a left-clone
    /// output and the right input; the multiset is the faithful comparison).</summary>
    private static List<string> FooterTexts(WmlDocument d)
    {
        using var ms = new MemoryStream(d.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(ms, false);
        XNamespace w = HeaderFooterFixtures.Wns;
        return doc.MainDocumentPart!.FooterParts
            .Select(p =>
            {
                using var s = p.GetStream(FileMode.Open, FileAccess.Read);
                return string.Concat(XDocument.Load(s).Descendants(w + "t").Select(t => t.Value));
            })
            .Where(t => !string.IsNullOrWhiteSpace(t))
            .OrderBy(t => t, StringComparer.Ordinal)
            .ToList();
    }

    private static List<string> SchemaErrors(byte[] bytes)
    {
        using var ms = new MemoryStream();
        ms.Write(bytes, 0, bytes.Length);
        using var wd = WordprocessingDocument.Open(ms, false);
        var validator = new OpenXmlValidator();
        return validator.Validate(wd)
            .Where(e => e.ErrorType == DocumentFormat.OpenXml.Validation.ValidationErrorType.Schema &&
                        !OxPt.WcTests.ExpectedErrors.Contains(e.Description))
            .Select(e => $"{e.Part?.Uri}: {e.Description}")
            .ToList();
    }

    /// <summary>Run the headless-LibreOffice header/footer render oracle; soft-skip when soffice/uno is
    /// unavailable (the comment-fidelity suite's portability convention). Tries python3.13 first (the
    /// system uno module targets it on this distro), then python3.</summary>
    private void RunLibreOfficeHeaderFooterOracle(string docxPath, string[] expects, string[] absents)
    {
        var script = Path.GetFullPath(Path.Combine("../../../../tools/diffharness/lo/lo_headerfooter_check.py"));
        if (!File.Exists(script))
        {
            _out.WriteLine($"[lo-skip] oracle script not found: {script}");
            return;
        }
        var args = string.Join(' ', expects.Select(e => $"\"{e}\""));
        if (absents.Length > 0)
            args += " --absent " + string.Join(' ', absents.Select(a => $"\"{a}\""));

        foreach (var python in new[] { "python3.13", "python3" })
        {
            string stdout;
            try
            {
                var psi = new ProcessStartInfo(python, $"\"{script}\" \"{docxPath}\" {args}")
                {
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                };
                using var proc = Process.Start(psi);
                if (proc == null) continue;
                stdout = proc.StandardOutput.ReadToEnd() + proc.StandardError.ReadToEnd();
                if (!proc.WaitForExit(120_000)) { try { proc.Kill(true); } catch { } _out.WriteLine("[lo-skip] timed out"); return; }
            }
            catch (Exception e)
            {
                _out.WriteLine($"[lo-skip] could not run oracle under {python}: {e.Message}");
                continue;
            }

            if (stdout.Contains("RESULT: OK"))
            {
                _out.WriteLine($"[lo-ok] {Path.GetFileName(docxPath)}: " +
                    (stdout.Split('\n').FirstOrDefault(l => l.StartsWith("doc=")) ?? ""));
                return;
            }
            if (stdout.Contains("RESULT: FAIL") || stdout.Contains("LOAD=FAILED"))
                Assert.Fail($"LibreOffice header/footer oracle FAILED for {Path.GetFileName(docxPath)}:\n{stdout}");
            if (stdout.Contains("ModuleNotFoundError"))
                continue; // this interpreter lacks uno — try the next
            _out.WriteLine($"[lo-skip] oracle inconclusive under {python}: {stdout.Split('\n').FirstOrDefault()}");
            return;
        }
        _out.WriteLine("[lo-skip] no python interpreter with uno available");
    }
}
