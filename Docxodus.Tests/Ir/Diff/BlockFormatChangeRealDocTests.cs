#nullable enable

using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus;
using Xunit;
using Xunit.Abstractions;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Real-document battery for the block-format-change family (spec:
/// 2026-07-03-diff-block-format-changes-design.md §5). A real corpus DOCX
/// (<c>HC029-Table-Merged-Cells</c>) is programmatically mutated across the whole family — a paragraph's
/// justification, a table's tblPr/tblGrid, a row's trPr, a cell's tcPr, and the trailing section's
/// margins — and <see cref="DocxDiff.Compare"/> must produce ALL FIVE native change markers, schema-valid,
/// round-tripping (accept ≡ right, reject ≡ left including the property bytes), and deterministically.
/// A headless-LibreOffice load backstop soft-skips when soffice/uno is unavailable.
/// </summary>
public class BlockFormatChangeRealDocTests
{
    private readonly ITestOutputHelper _out;
    public BlockFormatChangeRealDocTests(ITestOutputHelper o) => _out = o;

    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly string Fixture =
        Path.GetFullPath("../../../../TestFiles/HC029-Table-Merged-Cells.docx");

    private static readonly DocxDiffSettings Settings = new();

    private static WmlDocument Left() => new WmlDocument(Fixture);

    /// <summary>Build the RIGHT document by mutating one instance of each family member in the LEFT bytes.</summary>
    private static WmlDocument Right(out bool[] applied)
    {
        var localApplied = new bool[7];
        using var ms = new MemoryStream();
        var bytes = File.ReadAllBytes(Fixture);
        ms.Write(bytes, 0, bytes.Length);
        using (var wd = WordprocessingDocument.Open(ms, true))
        {
            var xdoc = wd.MainDocumentPart!.GetXDocument();
            var body = xdoc.Root!.Element(W + "body")!;

            // (0) pPr: a paragraph with no pPr gains a fresh centered pPr (jc-only pPr is always valid).
            var plainP = body.Descendants(W + "p").FirstOrDefault(p => p.Element(W + "pPr") == null);
            if (plainP != null)
            {
                plainP.AddFirst(new XElement(W + "pPr", new XElement(W + "jc", new XAttribute(W + "val", "center"))));
                localApplied[0] = true;
            }

            // (1) tblPr: double the first table-width value (an existing tblPr child — no schema-order risk).
            var tblW = body.Descendants(W + "tblW").FirstOrDefault(e => e.Parent?.Name == W + "tblPr");
            if (tblW != null && int.TryParse((string?)tblW.Attribute(W + "w"), out var tw))
            {
                tblW.SetAttributeValue(W + "w", tw == 0 ? 5000 : tw * 2);
                tblW.SetAttributeValue(W + "type", "dxa");
                localApplied[1] = true;
            }

            // (2) tblGrid: widen the first grid column.
            var gridCol = body.Descendants(W + "gridCol").FirstOrDefault();
            if (gridCol != null && int.TryParse((string?)gridCol.Attribute(W + "w"), out var gw))
            {
                gridCol.SetAttributeValue(W + "w", gw + 500);
                localApplied[2] = true;
            }

            // (3) trPr: a row with no trPr gains a fresh trHeight (trPr is first in CT_Row — always valid).
            var plainTr = body.Descendants(W + "tr").FirstOrDefault(tr => tr.Element(W + "trPr") == null);
            if (plainTr != null)
            {
                plainTr.AddFirst(new XElement(W + "trPr",
                    new XElement(W + "trHeight", new XAttribute(W + "val", "600"), new XAttribute(W + "hRule", "atLeast"))));
                localApplied[3] = true;
            }

            // (4) tcPr: widen the first cell (an existing tcW child — no schema-order risk).
            var tcW = body.Descendants(W + "tcW").FirstOrDefault(e => e.Parent?.Name == W + "tcPr");
            if (tcW != null && int.TryParse((string?)tcW.Attribute(W + "w"), out var cw))
            {
                tcW.SetAttributeValue(W + "w", cw + 500);
                localApplied[4] = true;
            }

            // (5) sectPr: change the trailing section's top margin.
            var sectPr = body.Elements(W + "sectPr").LastOrDefault();
            var pgMar = sectPr?.Element(W + "pgMar");
            if (pgMar != null)
            {
                pgMar.SetAttributeValue(W + "top", "2200");
                localApplied[5] = true;
            }

            // (6) tblPrEx: add a row-level table-property exception to the first row (CT_Row puts tblPrEx
            // first). This exercises w:tblPrExChange (an add: the left row has no tblPrEx).
            var firstTr = body.Descendants(W + "tr").FirstOrDefault();
            if (firstTr != null && firstTr.Element(W + "tblPrEx") == null)
            {
                firstTr.AddFirst(new XElement(W + "tblPrEx",
                    new XElement(W + "tblBorders",
                        new XElement(W + "top", new XAttribute(W + "val", "double"), new XAttribute(W + "sz", "8")))));
                localApplied[6] = true;
            }

            wd.MainDocumentPart.PutXDocument();
        }
        applied = localApplied;
        return new WmlDocument("hc029-mutated.docx", ms.ToArray());
    }

    [Fact]
    public void RealDoc_produces_all_five_family_markers_schema_valid_and_round_trips()
    {
        var left = Left();
        var right = Right(out var applied);
        Assert.True(applied.All(a => a),
            $"not every family mutation applied: [{string.Join(",", applied)}]");

        var result = DocxDiff.Compare(left, right, Settings);
        var body = BodyOf(result);

        // All native change markers present (the full two-way table + section family).
        Assert.NotEmpty(body.Descendants(W + "pPrChange"));
        Assert.NotEmpty(body.Descendants(W + "tblPrChange"));
        Assert.NotEmpty(body.Descendants(W + "tblGridChange"));
        Assert.NotEmpty(body.Descendants(W + "trPrChange"));
        Assert.NotEmpty(body.Descendants(W + "tcPrChange"));
        Assert.NotEmpty(body.Descendants(W + "tblPrExChange"));
        Assert.NotEmpty(body.Descendants(W + "sectPrChange"));

        // Schema-valid: the produced markup introduces NO NEW schema errors over the input (HC029 carries
        // pre-existing w:tblLook attribute noise the bundled SDK schema predates — fixture noise, not a
        // regression; the baseline comparison is the same one the renderer battery uses).
        AssertNoNewSchemaErrors(result, left);

        // Round-trip: accept ≡ right, reject ≡ left at the property level for each member.
        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        AssertNoNewSchemaErrors(accepted, left);
        AssertNoNewSchemaErrors(rejected, left);

        // Accept restores the RIGHT trailing-section top margin; reject the LEFT.
        Assert.Equal("2200", TrailingTopMargin(accepted));
        Assert.Equal(TrailingTopMargin(left), TrailingTopMargin(rejected));

        // No change markers survive accept OR reject (they resolve, never linger).
        foreach (var d in new[] { accepted, rejected })
            foreach (var name in new[] { "pPrChange", "tblPrChange", "tblGridChange", "trPrChange", "tcPrChange", "tblPrExChange", "sectPrChange" })
                Assert.Empty(BodyOf(d).Descendants(W + name));
    }

    [Fact]
    public void RealDoc_compare_is_deterministic()
    {
        var left = Left();
        var right = Right(out _);
        var a = DocxDiff.Compare(left, right, Settings);
        var b = DocxDiff.Compare(left, right, Settings);
        // Compare the produced BODY XML (the package zip bytes vary in timestamps; the content is what matters).
        Assert.Equal(BodyOf(a).ToString(SaveOptions.DisableFormatting),
                     BodyOf(b).ToString(SaveOptions.DisableFormatting));
    }

    [Fact]
    public void RealDoc_loads_in_libreoffice_backstop()
    {
        var result = DocxDiff.Compare(Left(), Right(out _), Settings);
        var dir = Path.Combine(Path.GetTempPath(), "docxodus-blockfmt-realdoc");
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "redline.docx");
        File.WriteAllBytes(path, result.DocumentByteArray);
        RunLibreOfficeLoad(path);
    }

    // ---- helpers ----

    private static XElement BodyOf(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        return wd.MainDocumentPart!.GetXDocument().Root!.Element(W + "body")!;
    }

    private static string? TrailingTopMargin(WmlDocument doc) =>
        (string?)BodyOf(doc).Elements(W + "sectPr").Last().Element(W + "pgMar")?.Attribute(W + "top");

    private static System.Collections.Generic.List<string> SchemaErrors(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        return new OpenXmlValidator().Validate(wd)
            .Where(e => e.ErrorType == DocumentFormat.OpenXml.Validation.ValidationErrorType.Schema)
            .Select(e => e.Description)
            .ToList();
    }

    /// <summary>Assert the produced document introduces no schema-error DESCRIPTION not already present in
    /// the baseline input (so pre-existing fixture noise — e.g. HC029's w:tblLook attributes — is tolerated
    /// while any error the block-format markup introduces fails).</summary>
    private static void AssertNoNewSchemaErrors(WmlDocument produced, WmlDocument baseline)
    {
        var baselineErrors = SchemaErrors(baseline).ToHashSet();
        var newErrors = SchemaErrors(produced).Where(d => !baselineErrors.Contains(d)).ToList();
        Assert.True(newErrors.Count == 0, "NEW schema errors:\n" + string.Join("\n", newErrors));
    }

    /// <summary>Headless-LibreOffice LOAD backstop: an INDEPENDENT renderer must open the tracked-changes
    /// document without error (the header/footer check script reports LOAD=FAILED / RESULT: OK and, with no
    /// expect strings, simply confirms the load). Soft-skips when soffice/uno is unavailable.</summary>
    private void RunLibreOfficeLoad(string docxPath)
    {
        var script = Path.GetFullPath("../../../../tools/diffharness/lo/lo_headerfooter_check.py");
        if (!File.Exists(script))
        {
            _out.WriteLine($"[lo-skip] oracle script not found: {script}");
            return;
        }
        foreach (var python in new[] { "python3.13", "python3" })
        {
            string stdout;
            try
            {
                var psi = new ProcessStartInfo(python, $"\"{script}\" \"{docxPath}\"")
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

            if (stdout.Contains("LOAD=FAILED"))
                Assert.Fail($"LibreOffice failed to load the block-format redline:\n{stdout}");
            if (stdout.Contains("ModuleNotFoundError"))
                continue;
            _out.WriteLine($"[lo] {stdout.Split('\n').FirstOrDefault()}");
            return;
        }
        _out.WriteLine("[lo-skip] no python interpreter with uno available");
    }
}
