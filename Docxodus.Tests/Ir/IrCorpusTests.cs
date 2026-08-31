#nullable enable

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Ir;
using Xunit;
using Xunit.Abstractions;

namespace Docxodus.Tests.Ir;

/// <summary>
/// Conformance tests for the IR reader (spec §10): totality over the whole <c>TestFiles/</c>
/// corpus, plus golden-snapshot diagnostic JSON over a curated handful of fixtures.
/// </summary>
/// <remarks>
/// <para><b>Regenerating snapshots.</b> Set the environment variable
/// <c>DOCXODUS_IR_REGEN_SNAPSHOTS=1</c> and run
/// <c>Read_CuratedFixtures_MatchGoldenSnapshots</c>. Each case then writes its
/// <c>.ir.json</c> into the <em>source tree</em> (<c>Docxodus.Tests/Ir/Snapshots/</c>) instead of
/// asserting, and the test is marked skipped. Diff every regenerated snapshot under review — the
/// diagnostic JSON is not a versioned contract, but a snapshot diff is a real IR-behavior change
/// that must be triaged (spec §9, §11). Never blind-regenerate.</para>
/// </remarks>
public class IrCorpusTests
{
    private static readonly DirectoryInfo TestFilesDir = new("../../../../TestFiles/");

    private readonly ITestOutputHelper _output;

    public IrCorpusTests(ITestOutputHelper output) => _output = output;

    /// <summary>
    /// Issue #618: reading with <see cref="UnidAssignment.IdentityBearing"/> must produce a
    /// byte-identical IR to reading with <see cref="UnidAssignment.All"/>, over the whole corpus.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The pruned pass skips assigning a Unid to elements no consumer can address — run/table/row/cell
    /// property containers, page setup, paragraph properties other than an inline section break, and
    /// text leaves. That is safe only if the assigned VALUES do not move, and the reason they cannot
    /// is structural: a Unid derives from <c>parentUnid : tag : signature : dupIndex</c> — ancestors
    /// only, never descendants — and the skip predicate is keyed on element names, so it removes whole
    /// tag-groups at once and no surviving element's <c>dupIndex</c> can shift.
    /// </para>
    /// <para>
    /// This asserts that reasoning against real documents rather than trusting it. The diagnostic JSON
    /// is the right comparison surface because it carries EVERY identity the reader hands out — block,
    /// row, cell, section-break, note, comment and content-control anchors, the inline section-break
    /// anchor, and the <c>w:drawing</c> Unid on an inline image, which is deliberately equality-neutral
    /// on <c>IrInlineImage</c> and so would slip through a structural record comparison.
    /// </para>
    /// </remarks>
    [Fact]
    [Trait("Category", "Corpus")]
    public void Read_PrunedUnidAssignment_ProducesAnIdenticalIr()
    {
        var all = new IrReaderOptions { UnidAssignment = UnidAssignment.All };
        var pruned = new IrReaderOptions { UnidAssignment = UnidAssignment.IdentityBearing };

        var files = TestFilesDir.GetFiles("*.docx", SearchOption.AllDirectories)
            .OrderBy(f => f.FullName, StringComparer.Ordinal)
            .ToList();

        var divergent = new List<string>();
        int compared = 0;
        int withAnchors = 0;
        foreach (var file in files)
        {
            if (!CanOpen(file)) continue;

            string expected;
            string actual;
            try
            {
                expected = IrDiagnosticJson.Write(IrReader.Read(new WmlDocument(file.FullName), all));
                actual = IrDiagnosticJson.Write(IrReader.Read(new WmlDocument(file.FullName), pruned));
            }
            catch (Exception)
            {
                // Totality is Read_EntireTestFilesCorpus_DoesNotThrow's job, not this test's.
                continue;
            }

            compared++;
            if (expected.Contains("\"anchor\"", StringComparison.Ordinal)) withAnchors++;
            if (!string.Equals(expected, actual, StringComparison.Ordinal))
                divergent.Add(file.Name);
        }

        // Non-vacuous: the corpus really was read, and the JSON really does carry anchors.
        Assert.True(compared > 500, $"only {compared} documents compared");
        Assert.True(withAnchors > 500, $"only {withAnchors} documents produced anchors");
        Assert.True(divergent.Count == 0,
            $"{divergent.Count} document(s) read differently under the pruned Unid pass: "
            + string.Join(", ", divergent.Take(10)));
        _output.WriteLine($"identical IR on {compared} documents ({withAnchors} carrying anchors)");
    }

    /// <summary>
    /// Totality proof: the reader must consume every readable fixture in <c>TestFiles/</c> without
    /// throwing (opaque fallback is the contract; a crash on weird-but-valid OOXML is a reader bug).
    /// Files that even <see cref="WordprocessingDocument.Open(Stream, bool)"/> rejects are
    /// intentionally-broken/encrypted fixtures and are recorded as skipped.
    /// </summary>
    [Fact]
    [Trait("Category", "Corpus")]
    public void Read_EntireTestFilesCorpus_DoesNotThrow()
    {
        var files = TestFilesDir.GetFiles("*.docx", SearchOption.AllDirectories)
            .OrderBy(f => f.FullName, StringComparer.Ordinal)
            .ToList();

        var failures = new List<string>();
        var skipped = new List<string>();
        int processed = 0;
        var sw = Stopwatch.StartNew();

        foreach (var file in files)
        {
            // A file that won't even open read-only is an intentionally-broken/encrypted fixture.
            if (!CanOpen(file))
            {
                skipped.Add(file.Name);
                continue;
            }

            try
            {
                _ = IrReader.Read(new WmlDocument(file.FullName));
                processed++;
            }
            catch (Exception ex)
            {
                failures.Add($"{file.Name}: {ex.GetType().Name}: {ex.Message}");
            }
        }

        sw.Stop();
        _output.WriteLine($"Corpus: {files.Count} *.docx; processed {processed}, " +
                          $"skipped {skipped.Count}, failed {failures.Count} in {sw.ElapsedMilliseconds} ms.");
        if (skipped.Count > 0)
            _output.WriteLine("Skipped (unreadable as OOXML): " + string.Join(", ", skipped));

        Assert.True(processed > 0, "No fixtures were processed — is the TestFiles path correct?");
        Assert.True(failures.Count == 0,
            $"IrReader.Read threw on {failures.Count} fixture(s):{Environment.NewLine}" +
            string.Join(Environment.NewLine, failures));
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

    // --- golden snapshots -------------------------------------------------

    // Curated fixtures spanning the IR's interesting shapes. Each entry's comment is the rationale.
    public static IEnumerable<object[]> CuratedFixtures()
    {
        // Simple text paragraphs — the baseline paragraph/inline shape.
        yield return new object[] { "HC006-Test-01.docx" };
        // Tables — exercises rows/cells/gridSpan/vMerge and the table content/format rollup.
        yield return new object[] { "HC029-Table-Merged-Cells.docx" };
        // Tracked changes — read under default Accept; proves revision normalization (N13).
        yield return new object[] { "RA001-Tracked-Revisions-01.docx" };
        // Lists / numbering — exercises IrListInfo (numId/ilvl) on paragraphs.
        yield return new object[] { "DB012-Lists-With-Different-Numberings.docx" };
        // Images — drawings become opaque inlines; covers the opaque-inline path.
        yield return new object[] { "HC042-Image-Png.docx" };
        // Footnotes — note references in body runs (opaque inline in M1.1 body scope).
        yield return new object[] { "DB007-Notes.docx" };
        // Headers/footers present — the body still reads; section/sectPr coverage.
        yield return new object[] { "DB002-Sections-With-Headers.docx" };
        // Complex/large document — broad coverage of mixed content.
        yield return new object[] { "HC031-Complicated-Document.docx" };
        // Textbox (M1.5) — w:txbxContent body modeled as IrTextbox: inner paragraph anchored + hashed,
        // emitted twice (DrawingML mc:Choice + VML mc:Fallback) mirroring the oracle's both-copies walk.
        yield return new object[] { "WC044-Text-Box.docx" };
        // Table-in-textbox (M1.5) — an inner w:tbl inside a textbox body: exercises the recursive block
        // walk + index registration through a textbox containing a table (rows/cells/cell blocks).
        yield return new object[] { "WC050-Table-in-Text-Box.docx" };
    }

    [Theory]
    [MemberData(nameof(CuratedFixtures))]
    public void Read_CuratedFixtures_MatchGoldenSnapshots(string fixtureName)
    {
        // Fixtures live both at the TestFiles/ root and in subfolders (e.g. WC/); resolve by name.
        var file = File.Exists(Path.Combine(TestFilesDir.FullName, fixtureName))
            ? new FileInfo(Path.Combine(TestFilesDir.FullName, fixtureName))
            : TestFilesDir.GetFiles(fixtureName, SearchOption.AllDirectories)
                .OrderBy(f => f.FullName, StringComparer.Ordinal).First();
        var doc = new WmlDocument(file.FullName);
        var json = IrDiagnosticJson.Write(IrReader.Read(doc));

        var snapshotPath = Path.Combine(SnapshotsDir(), Path.GetFileNameWithoutExtension(fixtureName) + ".ir.json");

        if (Environment.GetEnvironmentVariable("DOCXODUS_IR_REGEN_SNAPSHOTS") == "1")
        {
            // Regeneration mode: write the snapshot into the source tree and pass without asserting.
            // Every regenerated file must be reviewed/triaged before committing (spec §9, §11).
            Directory.CreateDirectory(Path.GetDirectoryName(snapshotPath)!);
            File.WriteAllText(snapshotPath, json, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            _output.WriteLine($"Regenerated snapshot: {snapshotPath}");
            return;
        }

        Assert.True(File.Exists(snapshotPath),
            $"Missing golden snapshot '{snapshotPath}'. Regenerate with DOCXODUS_IR_REGEN_SNAPSHOTS=1.");
        var expected = File.ReadAllText(snapshotPath);
        Assert.Equal(expected, json);
    }

    /// <summary>
    /// The source-tree <c>Snapshots/</c> directory, resolved from this file's compile-time path so
    /// regeneration writes into the source tree (not <c>bin/</c>), matching how the harness reads
    /// it back at runtime via the relative <see cref="TestFilesDir"/>-style path.
    /// </summary>
    private static string SnapshotsDir([CallerFilePath] string thisFile = "") =>
        Path.Combine(Path.GetDirectoryName(thisFile)!, "Snapshots");
}
