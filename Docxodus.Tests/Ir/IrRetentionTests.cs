#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Ir;
using Xunit;
using Xunit.Abstractions;

namespace Docxodus.Tests.Ir;

/// <summary>
/// M1.5 Task 2: <see cref="IrReaderOptions.RetainSources"/> = <c>false</c> behavior. Reading with
/// retention off must (a) stay total over the whole corpus, (b) leave <see cref="IrDocument.Sources"/>
/// empty and every node's <see cref="IrProvenance.Element"/> null, and (c) produce identical
/// content-addressable facts (anchors, <c>ContentHash</c>, <c>FormatFingerprint</c>) to a retained
/// read of the same bytes — retention is a memory optimization, never a content change.
/// </summary>
public class IrRetentionTests
{
    private static readonly DirectoryInfo TestFilesDir = new("../../../../TestFiles/");

    private readonly ITestOutputHelper _output;

    public IrRetentionTests(ITestOutputHelper output) => _output = output;

    private static readonly IrReaderOptions RetentionOff = new() { RetainSources = false };

    /// <summary>
    /// Totality with retention off, over the same corpus the retained totality test sweeps. Also
    /// asserts the snapshot drops its source pins: <see cref="IrDocument.Sources"/> is empty and no
    /// body/header/footer/note/comment block carries a non-null provenance <c>Element</c>.
    /// </summary>
    [Fact]
    [Trait("Category", "Corpus")]
    public void Read_EntireCorpus_RetentionOff_TotalAndUnpinned()
    {
        var files = TestFilesDir.GetFiles("*.docx", SearchOption.AllDirectories)
            .OrderBy(f => f.FullName, StringComparer.Ordinal)
            .ToList();

        var failures = new List<string>();
        int processed = 0;

        foreach (var file in files)
        {
            if (!CanOpen(file))
                continue;

            try
            {
                var ir = IrReader.Read(new WmlDocument(file.FullName), RetentionOff);

                if (ir.Sources.Count != 0)
                    failures.Add($"{file.Name}: Sources not empty ({ir.Sources.Count} entries)");

                var pinned = AllBlocks(ir).FirstOrDefault(b => b.Source.Element is not null);
                if (pinned is not null)
                    failures.Add($"{file.Name}: a {pinned.GetType().Name} retained a provenance Element");

                processed++;
            }
            catch (Exception ex)
            {
                failures.Add($"{file.Name}: {ex.GetType().Name}: {ex.Message}");
            }
        }

        _output.WriteLine($"Retention-off corpus: processed {processed} of {files.Count} *.docx.");

        Assert.True(processed > 0, "No fixtures were processed — is the TestFiles path correct?");
        Assert.True(failures.Count == 0,
            $"Retention-off read failed on {failures.Count} fixture(s):{Environment.NewLine}" +
            string.Join(Environment.NewLine, failures));
    }

    /// <summary>
    /// Spot-equality across a handful of structurally varied fixtures: a retained and a retention-off
    /// read of the SAME bytes must agree on every body block's anchor, <c>ContentHash</c>, and
    /// <c>FormatFingerprint</c> (the content-addressable facts the diff engine consumes). Provenance is
    /// the ONLY thing that differs.
    /// </summary>
    [Theory]
    [InlineData("HC031-Complicated-Document.docx")] // mixed content, the broadest body shape
    [InlineData("HC029-Table-Merged-Cells.docx")]   // tables → nested rows/cells/cell blocks
    [InlineData("DB012-Lists-With-Different-Numberings.docx")] // list numbering on paragraphs
    public void Read_RetentionOnVsOff_SameContentFacts(string fixtureName)
    {
        var file = ResolveFixture(fixtureName);
        var bytes = File.ReadAllBytes(file.FullName);

        var retained = IrReader.Read(new WmlDocument(fixtureName, (byte[])bytes.Clone()));
        var off = IrReader.Read(new WmlDocument(fixtureName, (byte[])bytes.Clone()), RetentionOff);

        var rb = retained.Body.Blocks.ToList();
        var ob = off.Body.Blocks.ToList();
        Assert.Equal(rb.Count, ob.Count);
        for (int i = 0; i < rb.Count; i++)
        {
            Assert.Equal(rb[i].Anchor.ToString(), ob[i].Anchor.ToString());
            Assert.Equal(rb[i].ContentHash.ToHex(), ob[i].ContentHash.ToHex());
            Assert.Equal(rb[i].FormatFingerprint.ToHex(), ob[i].FormatFingerprint.ToHex());
        }

        // The content scope compares value-equal (provenance is equality-neutral; the new scope-level
        // PartUri is identical for both reads since the source bytes/part are the same).
        Assert.Equal(retained.Body, off.Body);

        // Retention off drops the source pins; retention on keeps them.
        Assert.Empty(off.Sources);
        Assert.NotEmpty(retained.Sources);

        // The scope-level PartUri survives in BOTH modes (the promoted fact).
        Assert.NotNull(retained.Body.PartUri);
        Assert.Equal(retained.Body.PartUri, off.Body.PartUri);
    }

    /// <summary>
    /// Every block in every scope of a snapshot (body + headers/footers + notes + comments), with
    /// tables flattened to their cell blocks, for the unpinned-provenance sweep.
    /// </summary>
    private static IEnumerable<IrBlock> AllBlocks(IrDocument ir)
    {
        foreach (var b in Flatten(ir.Body.Blocks)) yield return b;
        foreach (var hf in ir.Headers.Concat(ir.Footers))
            foreach (var b in Flatten(hf.Scope.Blocks)) yield return b;
        foreach (var store in new[] { ir.Footnotes, ir.Endnotes })
            foreach (var note in store.Notes.Values)
                foreach (var b in Flatten(note.Blocks)) yield return b;
        foreach (var c in ir.Comments.Comments)
            foreach (var b in Flatten(c.Blocks)) yield return b;
    }

    private static IEnumerable<IrBlock> Flatten(IrNodeList<IrBlock> blocks)
    {
        foreach (var b in blocks)
        {
            yield return b;
            if (b is IrTable table)
                foreach (var row in table.Rows)
                    foreach (var cell in row.Cells)
                        foreach (var inner in Flatten(cell.Blocks))
                            yield return inner;
        }
    }

    private static FileInfo ResolveFixture(string fixtureName) =>
        File.Exists(Path.Combine(TestFilesDir.FullName, fixtureName))
            ? new FileInfo(Path.Combine(TestFilesDir.FullName, fixtureName))
            : TestFilesDir.GetFiles(fixtureName, SearchOption.AllDirectories)
                .OrderBy(f => f.FullName, StringComparer.Ordinal).First();

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
