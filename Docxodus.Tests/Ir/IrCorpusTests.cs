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
    }

    [Theory]
    [MemberData(nameof(CuratedFixtures))]
    public void Read_CuratedFixtures_MatchGoldenSnapshots(string fixtureName)
    {
        var doc = new WmlDocument(Path.Combine(TestFilesDir.FullName, fixtureName));
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
